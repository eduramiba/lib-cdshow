#include "stdafx.h"
#include "libcdshow.h"
#include "format_selection.h"

#pragma comment(lib, "strmiids.lib")

#include <windows.h>
#include <dshow.h>
#include <dvdmedia.h>
#include <strmif.h>
#include <comutil.h>
#include <comdef.h>

#include <atomic>
#include <string>
#include <thread>
#include <mutex>
#include <condition_variable>
#include <vector>
#include <cstdio>
#include <cinttypes>
#include <chrono>
#include <map>
#include <set>
#include <algorithm>
#include <limits>

#ifndef SAFE_RELEASE
#define SAFE_RELEASE(x) do { if ((x) != nullptr) { (x)->Release(); (x) = nullptr; } } while(0)
#endif

// ---- Sample Grabber & Null Renderer GUIDs (avoid qedit.h) ----
struct __declspec(uuid("C1F400A0-3F08-11d3-9F0B-006008039E37")) CLSID_SampleGrabber;
struct __declspec(uuid("6B652FFF-11FE-4fce-92AD-0266B5D7C78F")) IID_ISampleGrabber;
struct __declspec(uuid("C1F400A4-3F08-11d3-9F0B-006008039E37")) CLSID_NullRenderer;

MIDL_INTERFACE("0579154A-2B53-4994-B0D0-E773148EFF85")
ISampleGrabberCB : public IUnknown{
    virtual HRESULT STDMETHODCALLTYPE SampleCB(double, IMediaSample*) = 0;
    virtual HRESULT STDMETHODCALLTYPE BufferCB(double, BYTE*, long) = 0;
};

MIDL_INTERFACE("6B652FFF-11FE-4fce-92AD-0266B5D7C78F")
ISampleGrabber : public IUnknown{
    virtual HRESULT STDMETHODCALLTYPE SetOneShot(BOOL) = 0;
    virtual HRESULT STDMETHODCALLTYPE SetMediaType(const AM_MEDIA_TYPE*) = 0;
    virtual HRESULT STDMETHODCALLTYPE GetConnectedMediaType(AM_MEDIA_TYPE*) = 0;
    virtual HRESULT STDMETHODCALLTYPE SetBufferSamples(BOOL) = 0;
    virtual HRESULT STDMETHODCALLTYPE GetCurrentBuffer(long*, long*) = 0;
    virtual HRESULT STDMETHODCALLTYPE GetCurrentSample(IMediaSample**) = 0;
    virtual HRESULT STDMETHODCALLTYPE SetCallback(ISampleGrabberCB*, long) = 0;
};

// --------------------------- Logging helpers ---------------------------
static FILE* g_logFile = nullptr;
static std::once_flag g_logInitOnce;
static bool g_logEnabled = false;
static std::atomic<int> g_logOverride{ -1 }; // -1=use env/default, 0=off, 1=on

static bool parse_bool_env(const char* value) {
    if (!value || !*value) return false;
    std::string v(value);
    std::transform(v.begin(), v.end(), v.begin(), [](unsigned char c) { return (char)tolower(c); });
    if (v == "1" || v == "true" || v == "yes" || v == "on") return true;
    if (v == "0" || v == "false" || v == "no" || v == "off") return false;
    return false;
}

static bool is_debug_logging_enabled() {
    int ov = g_logOverride.load(std::memory_order_relaxed);
    if (ov >= 0) return ov != 0;

    std::call_once(g_logInitOnce, []() {
        // Disabled by default in all builds. Enable with: libcdshow_DEBUG=1
        char buf[32]{};
        DWORD n = GetEnvironmentVariableA("libcdshow_DEBUG", buf, (DWORD)sizeof(buf));
        if (n > 0 && n < sizeof(buf)) {
            g_logEnabled = parse_bool_env(buf);
        }
        else {
            g_logEnabled = false;
        }
    });
    ov = g_logOverride.load(std::memory_order_relaxed);
    if (ov >= 0) return ov != 0;
    return g_logEnabled;
}

static void dbg_print_raw(const char* s) {
    if (!is_debug_logging_enabled()) return;
    OutputDebugStringA(s);
    if (g_logFile) { fputs(s, g_logFile); fflush(g_logFile); }
    fputs(s, stderr); fflush(stderr);
}

static void dbg_printf(const char* fmt, ...) {
    if (!is_debug_logging_enabled()) return;
    char buf[2048];
    va_list ap; va_start(ap, fmt);
    _vsnprintf_s(buf, sizeof(buf), _TRUNCATE, fmt, ap);
    va_end(ap);
    dbg_print_raw(buf);
}

static std::wstring BstrToW(BSTR b) { return b ? std::wstring(b, SysStringLen(b)) : L""; }

static std::string GuidToStr(const GUID& g) {
    char s[64];
    _snprintf_s(s, sizeof(s), _TRUNCATE,
        "{%08lX-%04X-%04X-%02X%02X-%02X%02X%02X%02X%02X%02X}",
        g.Data1, g.Data2, g.Data3,
        g.Data4[0], g.Data4[1], g.Data4[2], g.Data4[3],
        g.Data4[4], g.Data4[5], g.Data4[6], g.Data4[7]);
    return s;
}

static const char* SubTypeName(const GUID& st) {
    if (st == MEDIASUBTYPE_YUY2) return "YUY2";
    if (st == MEDIASUBTYPE_MJPG) return "MJPG";
    if (st == MEDIASUBTYPE_RGB24) return "RGB24";
    if (st == MEDIASUBTYPE_NV12) return "NV12";
    if (st == MEDIASUBTYPE_RGB32) return "RGB32";
    if (st == MEDIASUBTYPE_ARGB32) return "ARGB32";
    return nullptr;
}

static int SubTypeProcessingRank(const GUID& st) {
    using namespace cds_format_selection;
    if (st == MEDIASUBTYPE_RGB32) return PROCESSING_RGB32;
    if (st == MEDIASUBTYPE_RGB24) return PROCESSING_RGB24;
    if (st == MEDIASUBTYPE_NV12) return PROCESSING_NV12;
    if (st == MEDIASUBTYPE_YUY2) return PROCESSING_YUY2;
    if (st == MEDIASUBTYPE_MJPG) return PROCESSING_MJPG;
    if (st == MEDIASUBTYPE_ARGB32) return PROCESSING_ARGB32;
    return PROCESSING_UNKNOWN;
}

static int PixelFormatFromSubtype(const GUID& subtype) {
    if (subtype == MEDIASUBTYPE_RGB32 || subtype == MEDIASUBTYPE_ARGB32) {
        return CDS_PIXEL_FORMAT_BGRA;
    }
    if (subtype == MEDIASUBTYPE_NV12) {
        return CDS_PIXEL_FORMAT_NV12;
    }
    if (subtype == MEDIASUBTYPE_YUY2) {
        return CDS_PIXEL_FORMAT_YUY2;
    }
    if (subtype == MEDIASUBTYPE_MJPG) {
        return CDS_PIXEL_FORMAT_MJPEG;
    }
    return CDS_PIXEL_FORMAT_UNKNOWN;
}

static const char* PixelFormatName(int pixelFormat) {
    switch (pixelFormat) {
    case CDS_PIXEL_FORMAT_BGRA:
        return "BGRA";
    case CDS_PIXEL_FORMAT_NV12:
        return "NV12";
    case CDS_PIXEL_FORMAT_YUY2:
        return "YUY2";
    case CDS_PIXEL_FORMAT_MJPEG:
        return "MJPEG";
    default:
        return "UNKNOWN";
    }
}

static std::string HResultToString(HRESULT hr) {
    _com_error err(hr);
    wchar_t const* msg = err.ErrorMessage();
    char mbs[512];
    WideCharToMultiByte(CP_UTF8, 0, msg, -1, mbs, sizeof(mbs), nullptr, nullptr);
    char buf[640];
    _snprintf_s(buf, sizeof(buf), _TRUNCATE, "hr=0x%08X (%s)", (unsigned)hr, mbs);
    return std::string(buf);
}

// ---- UTF-8 helpers ----
static std::string WStringToUtf8(const std::wstring& ws)
{
    if (ws.empty())
        return std::string();

    int sizeNeeded = WideCharToMultiByte(
        CP_UTF8,
        0,
        ws.c_str(),
        -1,              // include null terminator
        nullptr,
        0,
        nullptr,
        nullptr);

    if (sizeNeeded <= 0)
        return std::string();

    std::string result;
    result.resize((size_t)sizeNeeded);  // include null terminator

    int written = WideCharToMultiByte(
        CP_UTF8,
        0,
        ws.c_str(),
        -1,
        &result[0],      // <-- portable writable buffer
        sizeNeeded,
        nullptr,
        nullptr);
    if (written <= 0)
        return std::string();
    if (!result.empty() && result.back() == '\0')
        result.pop_back();
    else if ((size_t)written < result.size())
        result.resize((size_t)written);

    return result;
}

static std::string WToUtf8(const std::wstring& ws) {
    if (ws.empty()) return {};
    int sizeNeeded = WideCharToMultiByte(CP_UTF8, 0, ws.c_str(), -1, nullptr, 0, nullptr, nullptr);
    if (sizeNeeded <= 0) return {};
    std::string out;
    out.resize((size_t)sizeNeeded);
    int written = WideCharToMultiByte(CP_UTF8, 0, ws.c_str(), -1, &out[0], sizeNeeded, nullptr, nullptr);
    if (written <= 0) return {};
    if (!out.empty() && out.back() == '\0') out.pop_back();
    else if ((size_t)written < out.size()) out.resize((size_t)written);
    return out;
}

static std::string PinDirToStr(PIN_DIRECTION d) {
    return d == PINDIR_INPUT ? "IN" : "OUT";
}

// Try to read pin category via IKsPropertySet (AMPROPSETID_Pin / AMPROPERTY_PIN_CATEGORY)
static bool TryGetPinCategory(IPin* pin, GUID& outCat) {
    outCat = GUID{};
    IKsPropertySet* ks = nullptr;
    if (FAILED(pin->QueryInterface(IID_IKsPropertySet, (void**)&ks)) || !ks) return false;

    DWORD cb = 0;
    HRESULT hr = ks->Get(
        AMPROPSETID_Pin,
        AMPROPERTY_PIN_CATEGORY,
        nullptr, 0,
        &outCat, sizeof(outCat),
        &cb
    );
    ks->Release();
    return SUCCEEDED(hr) && cb == sizeof(GUID);
}

static std::wstring TryGetPinId(IPin* pin) {
    // IPin::QueryId returns allocated OLE string
    LPOLESTR id = nullptr;
    if (SUCCEEDED(pin->QueryId(&id)) && id) {
        std::wstring ws(id);
        CoTaskMemFree(id);
        return ws;
    }
    return L"";
}

static std::string CategoryName(const GUID& cat) {
    if (cat == PIN_CATEGORY_CAPTURE) return "PIN_CATEGORY_CAPTURE";
    if (cat == PIN_CATEGORY_PREVIEW) return "PIN_CATEGORY_PREVIEW";
    if (cat == PIN_CATEGORY_STILL) return "PIN_CATEGORY_STILL";
    if (cat == PIN_CATEGORY_ANALOGVIDEOIN) return "PIN_CATEGORY_ANALOGVIDEOIN";
    if (cat == PIN_CATEGORY_VBI) return "PIN_CATEGORY_VBI";
    if (cat == PIN_CATEGORY_CC) return "PIN_CATEGORY_CC";
    if (cat == PIN_CATEGORY_EDS) return "PIN_CATEGORY_EDS";
    if (cat == PIN_CATEGORY_TELETEXT) return "PIN_CATEGORY_TELETEXT";
    if (cat == PIN_CATEGORY_NABTS) return "PIN_CATEGORY_NABTS";
    // Otherwise GUID string:
    return GuidToStr(cat);
}

static void DumpFilterPins(IBaseFilter* filter, const char* tag)
{
    if (!filter) {
        dbg_printf("[%s] DumpFilterPins: filter=null\n", tag ? tag : "pins");
        return;
    }

    FILTER_INFO fi{};
    std::wstring fname = L"";
    if (SUCCEEDED(filter->QueryFilterInfo(&fi))) {
        fname = fi.achName;
        if (fi.pGraph) fi.pGraph->Release();
    }

    dbg_printf("========== PIN DUMP [%s] Filter='%s' ==========\n",
        tag ? tag : "pins",
        WToUtf8(fname).c_str()
    );

    IEnumPins* en = nullptr;
    HRESULT hr = filter->EnumPins(&en);
    if (FAILED(hr) || !en) {
        dbg_printf("EnumPins failed: %s\n", HResultToString(hr).c_str());
        return;
    }

    ULONG got = 0;
    IPin* pin = nullptr;
    int idx = 0;

    while (en->Next(1, &pin, &got) == S_OK && pin) {
        PIN_DIRECTION dir{};
        pin->QueryDirection(&dir);

        std::wstring pidW = TryGetPinId(pin);
        std::string pid = WToUtf8(pidW);

        GUID cat{};
        bool hasCat = TryGetPinCategory(pin, cat);

        // Connected?
        IPin* connectedTo = nullptr;
        bool connected = SUCCEEDED(pin->ConnectedTo(&connectedTo)) && connectedTo;

        dbg_printf("Pin #%d: dir=%s id='%s' category=%s connected=%s\n",
            idx,
            PinDirToStr(dir).c_str(),
            pid.c_str(),
            hasCat ? CategoryName(cat).c_str() : "(none)",
            connected ? "YES" : "NO"
        );

        if (connectedTo) connectedTo->Release();
        pin->Release();
        idx++;
    }

    en->Release();
    dbg_printf("========== END PIN DUMP [%s] ==========\n", tag ? tag : "pins");
}




static size_t copy_str(const std::string& s, char* buf, size_t len) {
    if (!buf || len == 0) return 0;
    size_t n = (s.size() < (len - 1)) ? s.size() : (len - 1);
    memcpy(buf, s.data(), n);
    buf[n] = 0;
    return n;
}

static void parse_vid_pid_from_path(const std::string& p, int& vid, int& pid) {
    vid = 0; pid = 0;
    std::string lower = p;
    for (auto& c : lower) c = (char)tolower((unsigned char)c);

    auto vpos = lower.find("vid_");
    auto ppos = lower.find("pid_");

    if (vpos != std::string::npos && vpos + 8 <= lower.size()) {
        vid = (int)strtol(lower.substr(vpos + 4, 4).c_str(), nullptr, 16);
    }
    if (ppos != std::string::npos && ppos + 8 <= lower.size()) {
        pid = (int)strtol(lower.substr(ppos + 4, 4).c_str(), nullptr, 16);
    }
}

// ======================= Common DirectShow helpers =======================

static HRESULT FindPinByCategory(IBaseFilter* f, const GUID& cat, PIN_DIRECTION dir, IPin** out) {
    *out = nullptr;
    IEnumPins* en = nullptr;
    HRESULT hr = f->EnumPins(&en);
    if (FAILED(hr)) return hr;

    IPin* p = nullptr; ULONG got = 0;
    while (en->Next(1, &p, &got) == S_OK) {
        PIN_DIRECTION d; p->QueryDirection(&d);
        if (d == dir) {
            IKsPropertySet* ks = nullptr;
            GUID pinCat{}; DWORD cb = 0;
            if (SUCCEEDED(p->QueryInterface(IID_IKsPropertySet, (void**)&ks))) {
                if (SUCCEEDED(ks->Get(AMPROPSETID_Pin, AMPROPERTY_PIN_CATEGORY,
                    nullptr, 0, &pinCat, sizeof(pinCat), &cb))) {
                    if (pinCat == cat) {
                        *out = p;
                        ks->Release();
                        en->Release();
                        return S_OK;
                    }
                }
                ks->Release();
            }
        }
        p->Release();
    }
    en->Release();
    return HRESULT_FROM_WIN32(ERROR_NOT_FOUND);
}

// =============================================================================
// ===================== NEW DirectShow Capture API (cds_*) ====================
// =============================================================================

struct DsFormat {
    uint32_t width = 0;
    uint32_t height = 0;
    uint32_t maxFps = 0;
    GUID subtype{};
    uint32_t streamCapsIndex = 0; // real IAMStreamConfig index
};

struct FormatKey {
    uint32_t w, h, fps;
    GUID st;
};

struct FormatKeyLess {
    bool operator()(const FormatKey& a, const FormatKey& b) const {
        if (a.w != b.w) return a.w < b.w;
        if (a.h != b.h) return a.h < b.h;
        if (a.fps != b.fps) return a.fps < b.fps;
        return memcmp(&a.st, &b.st, sizeof(GUID)) < 0;
    }
};

struct DsDevice {
    std::wstring nameW;
    std::wstring devicePathW;
    std::string  nameUtf8;
    std::string  devicePathUtf8;
    std::string  modelIdUtf8;
    int vid = 0;
    int pid = 0;
    std::wstring monikerDisplayNameW;
    std::vector<DsFormat> formats; // deduped + sorted
};

struct DsSession;
static uint64_t now_ts100ns_utc();
static bool calc_frame_layout_bytes(uint32_t width, uint32_t height, size_t& rowBytes, size_t& totalBytes);
static bool update_native_layout_from_sample_length(DsSession* session, size_t sampleBytes);
static HRESULT bind_moniker_by_display_name(const std::wstring& displayName, IMoniker** outMk);

struct DsFrameLayout {
    int pixelFormat = CDS_PIXEL_FORMAT_UNKNOWN;
    size_t dataBytes = 0;
    int planeCount = 0;
    size_t planeOffsets[2] = { 0, 0 };
    size_t planeStrides[2] = { 0, 0 };
};

enum class DsControlApi {
    VideoProcAmp,
    CameraControl
};

enum class DsControlRequestType {
    None,
    GetRange,
    Get,
    Set
};

struct DsControlPayload {
    long value = 0;
    long flags = 0;
    long minValue = 0;
    long maxValue = 0;
    long step = 0;
    long defaultValue = 0;
    long capabilities = 0;
};

class FrameGrabberCB : public ISampleGrabberCB {
public:
    FrameGrabberCB(DsSession* s) : _ref(1), _s(s) {}
    HRESULT STDMETHODCALLTYPE QueryInterface(REFIID riid, void** ppv) override {
        if (riid == __uuidof(IUnknown) || riid == __uuidof(ISampleGrabberCB)) {
            *ppv = static_cast<ISampleGrabberCB*>(this);
            AddRef();
            return S_OK;
        }
        *ppv = nullptr;
        return E_NOINTERFACE;
    }
    ULONG STDMETHODCALLTYPE AddRef() override { return (ULONG)InterlockedIncrement(&_ref); }
    ULONG STDMETHODCALLTYPE Release() override {
        ULONG r = (ULONG)InterlockedDecrement(&_ref);
        if (r == 0) delete this;
        return r;
    }
    HRESULT STDMETHODCALLTYPE SampleCB(double, IMediaSample*) override { return E_NOTIMPL; }
    HRESULT STDMETHODCALLTYPE BufferCB(double, BYTE* buffer, long len) override;

private:
    volatile LONG _ref;
    DsSession* _s;
};

class StillButtonCB : public ISampleGrabberCB {
public:
    StillButtonCB(DsSession* s) : _ref(1), _s(s) {}
    HRESULT STDMETHODCALLTYPE QueryInterface(REFIID riid, void** ppv) override {
        if (riid == __uuidof(IUnknown) || riid == __uuidof(ISampleGrabberCB)) {
            *ppv = static_cast<ISampleGrabberCB*>(this);
            AddRef();
            return S_OK;
        }
        *ppv = nullptr;
        return E_NOINTERFACE;
    }
    ULONG STDMETHODCALLTYPE AddRef() override { return (ULONG)InterlockedIncrement(&_ref); }
    ULONG STDMETHODCALLTYPE Release() override {
        ULONG r = (ULONG)InterlockedDecrement(&_ref);
        if (r == 0) delete this;
        return r;
    }
    HRESULT STDMETHODCALLTYPE SampleCB(double sampleTime, IMediaSample*) override;
    HRESULT STDMETHODCALLTYPE BufferCB(double, BYTE*, long) override { return E_NOTIMPL; }

private:
    volatile LONG _ref;
    DsSession* _s;
};

struct DsSession {
    uint32_t deviceIndex = 0;
    uint32_t width = 0;
    uint32_t height = 0;

    std::vector<uint8_t> lastFrame;
    std::mutex frameMutex;
    std::condition_variable frameCv;
    std::atomic<bool> hasFrame{ false };
    std::atomic<bool> pullFrameValid{ false };
    std::atomic<bool> pullFrameRequested{ false };
    uint64_t pullFrameSequence = 0;

    std::mutex deliveryMutex;
    std::condition_variable deliveryCv;
    cds_frame_callback_t frameCallback = nullptr;
    void* frameCallbackUserData = nullptr;
    bool frameCallbackExclusive = false;
    uint32_t activeFrameCallbacks = 0;

    bool bottomUp = false;
    int outputMode = CDS_OUTPUT_BGRA;
    DsFrameLayout frameLayout{};

    // ---- Button (edge triggered) ----
    std::atomic<bool> buttonEdge{ false };
    std::atomic<uint64_t> lastButtonTs100ns{ 0 };

    // ---- IAMVideoControl trigger detection ----
    IAMVideoControl* videoCtrl = nullptr;      // thread-owned
    IPin* stillPinVC = nullptr;                // thread-owned
    long vcCaps = 0;
    long lastVcMode = 0;
    bool vcHasTrigger = false;
    bool useStillFallback = false;

    // ---- Camera property control ----
    IAMVideoProcAmp* videoProcAmp = nullptr;   // thread-owned
    IAMCameraControl* cameraControl = nullptr; // thread-owned
    std::mutex controlMutex;
    std::condition_variable controlCv;
    bool controlRequestPending = false;
    bool controlRequestCompleted = false;
    DsControlApi controlApi = DsControlApi::VideoProcAmp;
    DsControlRequestType controlRequestType = DsControlRequestType::None;
    long controlProperty = 0;
    DsControlPayload controlPayload{};
    HRESULT controlRequestHr = E_FAIL;

    std::atomic<bool> stopRequested{ false };
    std::thread worker;
    std::mutex startMutex;
    std::condition_variable startCv;
    bool startCompleted = false;
    cds_result_t startResult = CDS_ERR_UNKNOWN;

    IGraphBuilder* graph = nullptr;
    ICaptureGraphBuilder2* cap = nullptr;
    IBaseFilter* capFilter = nullptr;

    IBaseFilter* grabberFilter = nullptr;
    ISampleGrabber* grabber = nullptr;
    IBaseFilter* nullRenderer = nullptr;
    IMediaControl* mc = nullptr;
    IMediaEvent* me = nullptr;
    FrameGrabberCB* frameCbObj = nullptr;
    IBaseFilter* stillGrabberFilter = nullptr;
    ISampleGrabber* stillGrabber = nullptr;
    IBaseFilter* stillNullRenderer = nullptr;
    StillButtonCB* stillCbObj = nullptr;

    void release_graph_thread_only() {
        if (mc) mc->Stop();
        if (grabber) grabber->SetCallback(nullptr, 0);
        if (stillGrabber) stillGrabber->SetCallback(nullptr, 0);

        SAFE_RELEASE(me);
        SAFE_RELEASE(mc);

        SAFE_RELEASE(nullRenderer);
        SAFE_RELEASE(grabber);
        SAFE_RELEASE(grabberFilter);
        SAFE_RELEASE(stillNullRenderer);
        SAFE_RELEASE(stillGrabber);
        SAFE_RELEASE(stillGrabberFilter);

        SAFE_RELEASE(stillPinVC);
        SAFE_RELEASE(videoCtrl);
        SAFE_RELEASE(videoProcAmp);
        SAFE_RELEASE(cameraControl);

        SAFE_RELEASE(capFilter);
        SAFE_RELEASE(cap);
        SAFE_RELEASE(graph);

        if (frameCbObj) { frameCbObj->Release(); frameCbObj = nullptr; }
        if (stillCbObj) { stillCbObj->Release(); stillCbObj = nullptr; }
    }
};

HRESULT STDMETHODCALLTYPE StillButtonCB::SampleCB(double sampleTime, IMediaSample*) {
    if (!_s) return S_OK;
    uint64_t t = now_ts100ns_utc();
    _s->lastButtonTs100ns.store(t);
    _s->buttonEdge.store(true);
    dbg_printf("[STILL FALLBACK] button sample\n");
    return S_OK;
}

HRESULT STDMETHODCALLTYPE FrameGrabberCB::BufferCB(double, BYTE* buffer, long len) {
    if (!_s || !buffer || len <= 0) return S_OK;

    const size_t sampleBytes = (size_t)len;
    size_t expected = 0;
    bool bottomUp = false;

    {
        std::lock_guard<std::mutex> lk(_s->frameMutex);
        if (_s->outputMode == CDS_OUTPUT_NATIVE) {
            if (!update_native_layout_from_sample_length(_s, sampleBytes)) return S_OK;
            expected = _s->frameLayout.dataBytes;
        }
        else {
            size_t rowBytes = 0;
            if (!calc_frame_layout_bytes(_s->width, _s->height, rowBytes, expected)) return S_OK;
            bottomUp = _s->bottomUp;
        }
    }
    if (expected == 0 || sampleBytes < expected) return S_OK;

    cds_frame_callback_t callback = nullptr;
    void* callbackUserData = nullptr;
    bool callbackExclusive = false;
    {
        std::lock_guard<std::mutex> lk(_s->deliveryMutex);
        callback = _s->frameCallback;
        callbackUserData = _s->frameCallbackUserData;
        callbackExclusive = callback && _s->frameCallbackExclusive;
        if (callback) {
            ++_s->activeFrameCallbacks;
        }
    }

    if (callback) {
        callback(
            _s->deviceIndex,
            buffer,
            expected,
            bottomUp ? 1 : 0,
            callbackUserData);
        {
            std::lock_guard<std::mutex> lk(_s->deliveryMutex);
            if (_s->activeFrameCallbacks > 0) {
                --_s->activeFrameCallbacks;
            }
        }
        _s->deliveryCv.notify_all();
    }

    const bool copyForPull =
        !callbackExclusive ||
        _s->pullFrameRequested.exchange(false, std::memory_order_acq_rel);
    if (copyForPull) {
        std::lock_guard<std::mutex> lk(_s->frameMutex);
        _s->lastFrame.resize(expected);
        if (!bottomUp) {
            memcpy(_s->lastFrame.data(), buffer, expected);
        }
        else {
            const size_t rowBytes = _s->frameLayout.planeStrides[0];
            for (uint32_t y = 0; y < _s->height; ++y) {
                const uint8_t* srcRow = buffer + ((size_t)(_s->height - 1 - y) * rowBytes);
                uint8_t* dstRow = _s->lastFrame.data() + ((size_t)y * rowBytes);
                memcpy(dstRow, srcRow, rowBytes);
            }
        }
        ++_s->pullFrameSequence;
        _s->pullFrameValid.store(true, std::memory_order_release);
        _s->frameCv.notify_all();
    }

    _s->hasFrame.store(true, std::memory_order_release);
    return S_OK;
}

// ---- Global capture state ----
static std::mutex g_dsMutex;
static bool g_dsInitialized = false;
static uint64_t g_dsGeneration = 1;
static std::vector<DsDevice> g_dsDevices;
static std::map<uint32_t, DsSession*> g_dsSessions;

static bool try_get_vih_dimensions(const VIDEOINFOHEADER* vih, uint32_t& width, uint32_t& height) {
    if (!vih) return false;
    LONG w = vih->bmiHeader.biWidth;
    LONG h = vih->bmiHeader.biHeight;
    if (w <= 0 || h == 0 || h == (std::numeric_limits<LONG>::min)()) return false;

    width = (uint32_t)w;
    height = (uint32_t)(h < 0 ? -h : h);
    return true;
}

static bool try_get_vih2_dimensions(const VIDEOINFOHEADER2* vih, uint32_t& width, uint32_t& height) {
    if (!vih) return false;
    LONG w = vih->bmiHeader.biWidth;
    LONG h = vih->bmiHeader.biHeight;
    if (w <= 0 || h == 0 || h == (std::numeric_limits<LONG>::min)()) return false;

    width = (uint32_t)w;
    height = (uint32_t)(h < 0 ? -h : h);
    return true;
}

static bool try_get_media_type_dimensions(
    const AM_MEDIA_TYPE* mt,
    uint32_t& width,
    uint32_t& height)
{
    if (!mt || !mt->pbFormat) return false;
    if (mt->formattype == FORMAT_VideoInfo && mt->cbFormat >= sizeof(VIDEOINFOHEADER)) {
        return try_get_vih_dimensions((const VIDEOINFOHEADER*)mt->pbFormat, width, height);
    }
    if (mt->formattype == FORMAT_VideoInfo2 && mt->cbFormat >= sizeof(VIDEOINFOHEADER2)) {
        return try_get_vih2_dimensions((const VIDEOINFOHEADER2*)mt->pbFormat, width, height);
    }
    return false;
}

static REFERENCE_TIME media_type_frame_interval(const AM_MEDIA_TYPE* mt) {
    if (!mt || !mt->pbFormat) return 0;
    if (mt->formattype == FORMAT_VideoInfo && mt->cbFormat >= sizeof(VIDEOINFOHEADER)) {
        return ((const VIDEOINFOHEADER*)mt->pbFormat)->AvgTimePerFrame;
    }
    if (mt->formattype == FORMAT_VideoInfo2 && mt->cbFormat >= sizeof(VIDEOINFOHEADER2)) {
        return ((const VIDEOINFOHEADER2*)mt->pbFormat)->AvgTimePerFrame;
    }
    return 0;
}

static bool set_media_type_frame_interval(AM_MEDIA_TYPE* mt, REFERENCE_TIME interval) {
    if (!mt || !mt->pbFormat || interval <= 0) return false;
    if (mt->formattype == FORMAT_VideoInfo && mt->cbFormat >= sizeof(VIDEOINFOHEADER)) {
        ((VIDEOINFOHEADER*)mt->pbFormat)->AvgTimePerFrame = interval;
        return true;
    }
    if (mt->formattype == FORMAT_VideoInfo2 && mt->cbFormat >= sizeof(VIDEOINFOHEADER2)) {
        ((VIDEOINFOHEADER2*)mt->pbFormat)->AvgTimePerFrame = interval;
        return true;
    }
    return false;
}

static bool calc_frame_layout_bytes(uint32_t width, uint32_t height, size_t& rowBytes, size_t& totalBytes) {
    if (width == 0 || height == 0) return false;
    if ((size_t)width > (SIZE_MAX / 4)) return false;
    rowBytes = (size_t)width * 4;
    if ((size_t)height > (SIZE_MAX / rowBytes)) return false;
    totalBytes = rowBytes * (size_t)height;
    return true;
}

static const BITMAPINFOHEADER* media_type_bitmap_header(const AM_MEDIA_TYPE* mt) {
    if (!mt || !mt->pbFormat) return nullptr;
    if (mt->formattype == FORMAT_VideoInfo && mt->cbFormat >= sizeof(VIDEOINFOHEADER)) {
        return &((const VIDEOINFOHEADER*)mt->pbFormat)->bmiHeader;
    }
    if (mt->formattype == FORMAT_VideoInfo2 && mt->cbFormat >= sizeof(VIDEOINFOHEADER2)) {
        return &((const VIDEOINFOHEADER2*)mt->pbFormat)->bmiHeader;
    }
    return nullptr;
}

static bool calc_connected_frame_layout(
    const AM_MEDIA_TYPE* mt,
    uint32_t width,
    uint32_t height,
    DsFrameLayout& layout)
{
    layout = DsFrameLayout{};
    if (!mt || width == 0 || height == 0) return false;

    layout.pixelFormat = PixelFormatFromSubtype(mt->subtype);
    const BITMAPINFOHEADER* bmi = media_type_bitmap_header(mt);
    const size_t advertisedBytes = bmi ? (size_t)bmi->biSizeImage : 0;

    if (layout.pixelFormat == CDS_PIXEL_FORMAT_BGRA) {
        size_t rowBytes = 0;
        size_t totalBytes = 0;
        if (!calc_frame_layout_bytes(width, height, rowBytes, totalBytes)) return false;
        layout.dataBytes = totalBytes;
        layout.planeCount = 1;
        layout.planeStrides[0] = rowBytes;
        return true;
    }

    if (layout.pixelFormat == CDS_PIXEL_FORMAT_YUY2) {
        if ((width & 1u) != 0 || (size_t)width > SIZE_MAX / 2u) return false;
        const size_t minimumStride = (size_t)width * 2u;
        size_t stride = minimumStride;
        if (advertisedBytes >= minimumStride * (size_t)height &&
            advertisedBytes % (size_t)height == 0) {
            stride = advertisedBytes / (size_t)height;
        }
        if ((size_t)height > SIZE_MAX / stride) return false;
        layout.dataBytes = stride * (size_t)height;
        layout.planeCount = 1;
        layout.planeStrides[0] = stride;
        return true;
    }

    if (layout.pixelFormat == CDS_PIXEL_FORMAT_NV12) {
        if ((width & 1u) != 0 || (height & 1u) != 0) return false;
        const size_t chromaRows = ((size_t)height + 1u) / 2u;
        const size_t totalRows = (size_t)height + chromaRows;
        size_t stride = (size_t)width;
        if (advertisedBytes >= stride * totalRows && advertisedBytes % totalRows == 0) {
            stride = advertisedBytes / totalRows;
        }
        if (totalRows > SIZE_MAX / stride) return false;
        layout.dataBytes = stride * totalRows;
        layout.planeCount = 2;
        layout.planeOffsets[1] = stride * (size_t)height;
        layout.planeStrides[0] = stride;
        layout.planeStrides[1] = stride;
        return true;
    }

    if (layout.pixelFormat == CDS_PIXEL_FORMAT_MJPEG) {
        layout.dataBytes = advertisedBytes;
        layout.planeCount = 1;
        // Compressed frames do not have a meaningful row stride.
        layout.planeStrides[0] = 0;
        return true;
    }

    return false;
}

static bool update_native_layout_from_sample_length(
    DsSession* session,
    size_t sampleBytes)
{
    if (!session || session->width == 0 || session->height == 0) return false;
    DsFrameLayout& layout = session->frameLayout;

    if (layout.pixelFormat == CDS_PIXEL_FORMAT_YUY2) {
        const size_t height = (size_t)session->height;
        const size_t minimumStride = (size_t)session->width * 2u;
        if (sampleBytes % height != 0) return sampleBytes >= layout.dataBytes;
        const size_t stride = sampleBytes / height;
        if (stride < minimumStride) return false;
        layout.dataBytes = sampleBytes;
        layout.planeCount = 1;
        layout.planeOffsets[0] = 0;
        layout.planeStrides[0] = stride;
        return true;
    }

    if (layout.pixelFormat == CDS_PIXEL_FORMAT_NV12) {
        const size_t chromaRows = ((size_t)session->height + 1u) / 2u;
        const size_t totalRows = (size_t)session->height + chromaRows;
        if (sampleBytes % totalRows != 0) return sampleBytes >= layout.dataBytes;
        const size_t stride = sampleBytes / totalRows;
        if (stride < (size_t)session->width) return false;
        layout.dataBytes = sampleBytes;
        layout.planeCount = 2;
        layout.planeOffsets[0] = 0;
        layout.planeOffsets[1] = stride * (size_t)session->height;
        layout.planeStrides[0] = stride;
        layout.planeStrides[1] = stride;
        return true;
    }

    if (layout.pixelFormat == CDS_PIXEL_FORMAT_MJPEG) {
        layout.dataBytes = sampleBytes;
        layout.planeCount = 1;
        layout.planeOffsets[0] = 0;
        layout.planeStrides[0] = 0;
        return sampleBytes > 0;
    }

    return sampleBytes >= layout.dataBytes;
}

static HRESULT co_initialize_for_api_call(bool& didInit) {
    didInit = false;
    HRESULT hr = CoInitializeEx(nullptr, COINIT_MULTITHREADED);
    if (SUCCEEDED(hr)) {
        didInit = true;
        return hr;
    }
    if (hr == RPC_E_CHANGED_MODE) {
        return S_OK;
    }
    return hr;
}

static cds_result_t control_hresult_to_result(HRESULT hr) {
    if (SUCCEEDED(hr)) return CDS_OK;
    if (hr == E_POINTER || hr == E_INVALIDARG) return CDS_ERR_INVALID_ARGUMENT;
    if (hr == E_NOINTERFACE ||
        hr == HRESULT_FROM_WIN32(ERROR_NOT_SUPPORTED) ||
        hr == HRESULT_FROM_WIN32(ERROR_NOT_FOUND)) {
        return CDS_ERR_CONTROL_NOT_SUPPORTED;
    }
    return CDS_ERR_CONTROL_IO;
}

static HRESULT bind_filter_by_device(const DsDevice& dev, IBaseFilter** outFilter) {
    if (!outFilter) return E_POINTER;
    *outFilter = nullptr;

    IMoniker* mk = nullptr;
    HRESULT hr = bind_moniker_by_display_name(dev.monikerDisplayNameW, &mk);
    if (FAILED(hr) || !mk) return FAILED(hr) ? hr : E_FAIL;

    hr = mk->BindToObject(nullptr, nullptr, IID_IBaseFilter, (void**)outFilter);
    mk->Release();
    return hr;
}

static HRESULT execute_control_request(
    DsControlApi api,
    IAMVideoProcAmp* videoProcAmp,
    IAMCameraControl* cameraControl,
    DsControlRequestType requestType,
    long property,
    DsControlPayload* payload)
{
    if (!payload) return E_POINTER;

    switch (api) {
    case DsControlApi::VideoProcAmp:
        if (!videoProcAmp) return E_NOINTERFACE;
        switch (requestType) {
        case DsControlRequestType::GetRange:
            return videoProcAmp->GetRange(
                property,
                &payload->minValue,
                &payload->maxValue,
                &payload->step,
                &payload->defaultValue,
                &payload->capabilities);
        case DsControlRequestType::Get:
            return videoProcAmp->Get(property, &payload->value, &payload->flags);
        case DsControlRequestType::Set:
            return videoProcAmp->Set(property, payload->value, payload->flags);
        default:
            return E_INVALIDARG;
        }

    case DsControlApi::CameraControl:
        if (!cameraControl) return E_NOINTERFACE;
        switch (requestType) {
        case DsControlRequestType::GetRange:
            return cameraControl->GetRange(
                property,
                &payload->minValue,
                &payload->maxValue,
                &payload->step,
                &payload->defaultValue,
                &payload->capabilities);
        case DsControlRequestType::Get:
            return cameraControl->Get(property, &payload->value, &payload->flags);
        case DsControlRequestType::Set:
            return cameraControl->Set(property, payload->value, payload->flags);
        default:
            return E_INVALIDARG;
        }
    }

    return E_INVALIDARG;
}

static cds_result_t execute_control_request_with_temporary_filter(
    const DsDevice& dev,
    DsControlApi api,
    DsControlRequestType requestType,
    long property,
    DsControlPayload* payload)
{
    if (!payload) return CDS_ERR_INVALID_ARGUMENT;

    bool didInit = false;
    HRESULT hr = co_initialize_for_api_call(didInit);
    if (FAILED(hr)) return CDS_ERR_CONTROL_IO;

    IBaseFilter* filter = nullptr;
    IAMVideoProcAmp* videoProcAmp = nullptr;
    IAMCameraControl* cameraControl = nullptr;

    hr = bind_filter_by_device(dev, &filter);
    if (SUCCEEDED(hr) && filter) {
        if (api == DsControlApi::VideoProcAmp) {
            hr = filter->QueryInterface(IID_IAMVideoProcAmp, (void**)&videoProcAmp);
        }
        else {
            hr = filter->QueryInterface(IID_IAMCameraControl, (void**)&cameraControl);
        }
    }

    if (SUCCEEDED(hr)) {
        hr = execute_control_request(api, videoProcAmp, cameraControl, requestType, property, payload);
    }

    SAFE_RELEASE(cameraControl);
    SAFE_RELEASE(videoProcAmp);
    SAFE_RELEASE(filter);
    if (didInit) CoUninitialize();

    return control_hresult_to_result(hr);
}

static void process_pending_control_request(DsSession* s) {
    if (!s) return;

    DsControlApi api = DsControlApi::VideoProcAmp;
    DsControlRequestType requestType = DsControlRequestType::None;
    long property = 0;
    DsControlPayload payload{};

    {
        std::unique_lock<std::mutex> lk(s->controlMutex);
        if (!s->controlRequestPending) return;
        api = s->controlApi;
        requestType = s->controlRequestType;
        property = s->controlProperty;
        payload = s->controlPayload;
        s->controlRequestPending = false;
    }

    HRESULT hr = execute_control_request(
        api,
        s->videoProcAmp,
        s->cameraControl,
        requestType,
        property,
        &payload);

    {
        std::lock_guard<std::mutex> lk(s->controlMutex);
        s->controlPayload = payload;
        s->controlRequestHr = hr;
        s->controlRequestCompleted = true;
        s->controlRequestType = DsControlRequestType::None;
    }
    s->controlCv.notify_all();
}

static void fail_pending_control_request(DsSession* s, HRESULT hr) {
    if (!s) return;

    bool notify = false;
    {
        std::lock_guard<std::mutex> lk(s->controlMutex);
        if (s->controlRequestPending || s->controlRequestType != DsControlRequestType::None) {
            s->controlRequestPending = false;
            s->controlRequestCompleted = true;
            s->controlRequestType = DsControlRequestType::None;
            s->controlRequestHr = hr;
            notify = true;
        }
    }

    if (notify) {
        s->controlCv.notify_all();
    }
}

static cds_result_t execute_control_request_with_active_session_locked(
    DsSession* s,
    DsControlApi api,
    DsControlRequestType requestType,
    long property,
    DsControlPayload* payload)
{
    if (!s || !payload) return CDS_ERR_INVALID_ARGUMENT;

    std::unique_lock<std::mutex> lk(s->controlMutex);
    s->controlApi = api;
    s->controlRequestType = requestType;
    s->controlProperty = property;
    s->controlPayload = *payload;
    s->controlRequestHr = E_FAIL;
    s->controlRequestCompleted = false;
    s->controlRequestPending = true;
    lk.unlock();
    s->controlCv.notify_all();

    lk.lock();
    s->controlCv.wait(lk, [&]() { return s->controlRequestCompleted; });
    *payload = s->controlPayload;
    return control_hresult_to_result(s->controlRequestHr);
}

static cds_result_t execute_device_control_request(
    int32_t device_index,
    DsControlApi api,
    DsControlRequestType requestType,
    long property,
    DsControlPayload* payload)
{
    std::lock_guard<std::mutex> lk(g_dsMutex);
    if (!g_dsInitialized) return CDS_ERR_NOT_INITIALIZED;
    if (device_index < 0 || (size_t)device_index >= g_dsDevices.size()) return CDS_ERR_DEVICE_NOT_FOUND;

    auto it = g_dsSessions.find((uint32_t)device_index);
    if (it != g_dsSessions.end()) {
        return execute_control_request_with_active_session_locked(it->second, api, requestType, property, payload);
    }

    return execute_control_request_with_temporary_filter(
        g_dsDevices[(size_t)device_index],
        api,
        requestType,
        property,
        payload);
}

// ---- Rebind by moniker display name ----
static HRESULT bind_moniker_by_display_name(const std::wstring& displayName, IMoniker** outMk) {
    *outMk = nullptr;
    IBindCtx* ctx = nullptr;
    HRESULT hr = CreateBindCtx(0, &ctx);
    if (FAILED(hr)) return hr;

    ULONG eaten = 0;
    IMoniker* mk = nullptr;
    hr = MkParseDisplayName(ctx, displayName.c_str(), &eaten, &mk);
    ctx->Release();
    if (FAILED(hr)) return hr;

    *outMk = mk;
    return S_OK;
}

static void free_am_media_type(AM_MEDIA_TYPE* mt) {
    if (!mt) return;
    if (mt->cbFormat && mt->pbFormat) CoTaskMemFree(mt->pbFormat);
    if (mt->pUnk) mt->pUnk->Release();
    CoTaskMemFree(mt);
}

static uint64_t now_ts100ns_utc() {
    FILETIME ft{};
    GetSystemTimeAsFileTime(&ft);
    ULARGE_INTEGER ui{};
    ui.LowPart = ft.dwLowDateTime;
    ui.HighPart = ft.dwHighDateTime;
    return (uint64_t)ui.QuadPart;
}

// ---- Enumerate devices + formats (dedup with correct mapping) ----
static HRESULT enumerate_devices_and_formats() {
    constexpr int kMaxStreamCapsBytes = 1024 * 1024;

    g_dsDevices.clear();

    ICreateDevEnum* devEnum = nullptr;
    IEnumMoniker* enumMon = nullptr;

    HRESULT hr = CoCreateInstance(CLSID_SystemDeviceEnum, nullptr, CLSCTX_INPROC_SERVER,
        IID_ICreateDevEnum, (void**)&devEnum);
    if (FAILED(hr)) return hr;

    hr = devEnum->CreateClassEnumerator(CLSID_VideoInputDeviceCategory, &enumMon, 0);
    devEnum->Release();
    // DirectShow returns S_FALSE when the category exists but currently has no
    // devices. That is a valid empty enumeration, not an initialization error.
    if (hr == S_FALSE) return S_OK;
    if (FAILED(hr)) return hr;

    IMoniker* mk = nullptr; ULONG fetched = 0;
    while (enumMon->Next(1, &mk, &fetched) == S_OK) {
        DsDevice dev{};

        // Read properties
        IPropertyBag* bag = nullptr;
        if (SUCCEEDED(mk->BindToStorage(nullptr, nullptr, IID_IPropertyBag, (void**)&bag))) {
            VARIANT vn; VariantInit(&vn);
            VARIANT vd; VariantInit(&vd);

            if (SUCCEEDED(bag->Read(L"FriendlyName", &vn, 0)) && vn.vt == VT_BSTR) {
                dev.nameW = BstrToW(vn.bstrVal);
            }
            if (SUCCEEDED(bag->Read(L"DevicePath", &vd, 0)) && vd.vt == VT_BSTR) {
                dev.devicePathW = BstrToW(vd.bstrVal);
            }

            VariantClear(&vn);
            VariantClear(&vd);
            bag->Release();
        }

        // Moniker display name for later binding
        {
            LPOLESTR dn = nullptr;
            if (SUCCEEDED(mk->GetDisplayName(nullptr, nullptr, &dn)) && dn) {
                dev.monikerDisplayNameW = dn;
                CoTaskMemFree(dn);
            }
        }

        dev.nameUtf8 = WStringToUtf8(dev.nameW);
        dev.devicePathUtf8 = WStringToUtf8(dev.devicePathW);
        dev.modelIdUtf8 = dev.nameUtf8;

        parse_vid_pid_from_path(dev.devicePathUtf8, dev.vid, dev.pid);

        // Enumerate formats using a temporary graph (robust)
        IBaseFilter* filter = nullptr;
        hr = mk->BindToObject(nullptr, nullptr, IID_IBaseFilter, (void**)&filter);
        if (SUCCEEDED(hr) && filter) {
            IGraphBuilder* graph = nullptr;
            ICaptureGraphBuilder2* cap = nullptr;

            HRESULT hrGraph = CoCreateInstance(CLSID_FilterGraph, nullptr, CLSCTX_INPROC_SERVER,
                IID_IGraphBuilder, (void**)&graph);
            HRESULT hrCap = SUCCEEDED(hrGraph)
                ? CoCreateInstance(CLSID_CaptureGraphBuilder2, nullptr, CLSCTX_INPROC_SERVER,
                    IID_ICaptureGraphBuilder2, (void**)&cap)
                : hrGraph;

            if (SUCCEEDED(hrGraph) && SUCCEEDED(hrCap)) {

                HRESULT hrFG = cap->SetFiltergraph(graph);
                if (SUCCEEDED(hrFG)) {
                    hrFG = graph->AddFilter(filter, L"Capture");
                }

                IAMStreamConfig* cfg = nullptr;
                HRESULT hrCfg = FAILED(hrFG) ? hrFG
                    : cap->FindInterface(&PIN_CATEGORY_CAPTURE, &MEDIATYPE_Video, filter, IID_IAMStreamConfig, (void**)&cfg);
                if (FAILED(hrCfg)) {
                    hrCfg = cap->FindInterface(&PIN_CATEGORY_PREVIEW, &MEDIATYPE_Video, filter, IID_IAMStreamConfig, (void**)&cfg);
                }

                // Dedup map keyed by (w,h,fps,subtype) and store first representative (streamCapsIndex)
                std::map<FormatKey, DsFormat, FormatKeyLess> uniq;

                if (SUCCEEDED(hrCfg) && cfg) {
                    int count = 0, size = 0;
                    HRESULT hrCaps = cfg->GetNumberOfCapabilities(&count, &size);
                    if (SUCCEEDED(hrCaps) &&
                        count > 0 &&
                        size >= (int)sizeof(VIDEO_STREAM_CONFIG_CAPS) &&
                        size <= kMaxStreamCapsBytes) {
                        std::vector<uint8_t> capsBuf((size_t)size);

                        for (int i = 0; i < count; ++i) {
                            AM_MEDIA_TYPE* mt = nullptr;
                            if (SUCCEEDED(cfg->GetStreamCaps(i, &mt, capsBuf.data())) && mt) {
                                VIDEO_STREAM_CONFIG_CAPS* caps = (VIDEO_STREAM_CONFIG_CAPS*)capsBuf.data();

                                uint32_t w = 0, h = 0;
                                if (!try_get_media_type_dimensions(mt, w, h)) {
                                    free_am_media_type(mt);
                                    continue;
                                }

                                uint32_t maxFps = 0;
                                if (caps->MinFrameInterval > 0) {
                                    maxFps = (uint32_t)(10000000ULL / (uint64_t)caps->MinFrameInterval);
                                }
                                else {
                                    REFERENCE_TIME interval = media_type_frame_interval(mt);
                                    if (interval > 0)
                                        maxFps = (uint32_t)(10000000ULL / (uint64_t)interval);
                                }

                                if (w && h) {
                                    FormatKey key{ w, h, maxFps, mt->subtype };
                                    if (uniq.find(key) == uniq.end()) {
                                        DsFormat f{};
                                        f.width = w;
                                        f.height = h;
                                        f.maxFps = maxFps;
                                        f.subtype = mt->subtype;
                                        f.streamCapsIndex = (uint32_t)i; // REAL index
                                        uniq.emplace(key, f);
                                    }
                                }

                                free_am_media_type(mt);
                            }
                        }
                    }
                    else {
                        dbg_printf("Invalid stream capabilities: hr=%s count=%d size=%d\n",
                            HResultToString(hrCaps).c_str(), count, size);
                    }

                    SAFE_RELEASE(cfg);
                }

                // Emit formats in stable sorted order (map iteration sorted by key)
                dev.formats.clear();
                dev.formats.reserve(uniq.size());
                for (auto& kv : uniq) {
                    dev.formats.push_back(kv.second);
                }
            }
            SAFE_RELEASE(cap);
            SAFE_RELEASE(graph);

            SAFE_RELEASE(filter);
        }

        g_dsDevices.push_back(std::move(dev));
        mk->Release();
    }

    enumMon->Release();
    return S_OK;
}

static void cleanup_still_fallback_branch(DsSession* s) {
    if (!s) return;
    if (s->stillGrabber) {
        s->stillGrabber->SetCallback(nullptr, 0);
    }
    if (s->graph && s->stillNullRenderer) {
        s->graph->RemoveFilter(s->stillNullRenderer);
    }
    if (s->graph && s->stillGrabberFilter) {
        s->graph->RemoveFilter(s->stillGrabberFilter);
    }
    SAFE_RELEASE(s->stillNullRenderer);
    SAFE_RELEASE(s->stillGrabber);
    SAFE_RELEASE(s->stillGrabberFilter);
    if (s->stillCbObj) { s->stillCbObj->Release(); s->stillCbObj = nullptr; }
}

static HRESULT build_still_fallback_button_branch(DsSession* s) {
    if (!s || !s->graph || !s->cap || !s->capFilter) return E_POINTER;

    cleanup_still_fallback_branch(s);

    HRESULT hrStill = CoCreateInstance(__uuidof(CLSID_SampleGrabber), nullptr, CLSCTX_INPROC_SERVER,
        IID_IBaseFilter, (void**)&s->stillGrabberFilter);
    dbg_printf("Fallback Create STILL SampleGrabber => %s\n", HResultToString(hrStill).c_str());
    if (FAILED(hrStill)) return hrStill;

    hrStill = s->graph->AddFilter(s->stillGrabberFilter, L"StillGrabber");
    dbg_printf("Fallback AddFilter(StillGrabber) => %s\n", HResultToString(hrStill).c_str());
    if (FAILED(hrStill)) { cleanup_still_fallback_branch(s); return hrStill; }

    hrStill = s->stillGrabberFilter->QueryInterface(__uuidof(ISampleGrabber), (void**)&s->stillGrabber);
    dbg_printf("Fallback QI(ISampleGrabber still) => %s\n", HResultToString(hrStill).c_str());
    if (FAILED(hrStill) || !s->stillGrabber) { cleanup_still_fallback_branch(s); return FAILED(hrStill) ? hrStill : E_FAIL; }

    s->stillGrabber->SetOneShot(FALSE);
    s->stillGrabber->SetBufferSamples(FALSE);

    hrStill = CoCreateInstance(__uuidof(CLSID_NullRenderer), nullptr, CLSCTX_INPROC_SERVER,
        IID_IBaseFilter, (void**)&s->stillNullRenderer);
    dbg_printf("Fallback Create StillNull => %s\n", HResultToString(hrStill).c_str());
    if (FAILED(hrStill)) { cleanup_still_fallback_branch(s); return hrStill; }

    hrStill = s->graph->AddFilter(s->stillNullRenderer, L"StillNull");
    dbg_printf("Fallback AddFilter(StillNull) => %s\n", HResultToString(hrStill).c_str());
    if (FAILED(hrStill)) { cleanup_still_fallback_branch(s); return hrStill; }

    IPin* stillOut = nullptr;
    hrStill = FindPinByCategory(s->capFilter, PIN_CATEGORY_STILL, PINDIR_OUTPUT, &stillOut);
    dbg_printf("Fallback FindPinByCategory(STILL) => %s\n", HResultToString(hrStill).c_str());
    if (FAILED(hrStill) || !stillOut) {
        cleanup_still_fallback_branch(s);
        return FAILED(hrStill) ? hrStill : E_FAIL;
    }

    // Match the legacy working path: force grabber MT to first STILL media type.
    IEnumMediaTypes* emt = nullptr;
    HRESULT hrE = stillOut->EnumMediaTypes(&emt);
    dbg_printf("Fallback EnumMediaTypes(STILL) => %s\n", HResultToString(hrE).c_str());
    if (SUCCEEDED(hrE) && emt) {
        AM_MEDIA_TYPE* stillMt = nullptr;
        HRESULT hrN = emt->Next(1, &stillMt, nullptr);
        dbg_printf("Fallback Get first STILL MT => %s\n", HResultToString(hrN).c_str());
        if (hrN == S_OK && stillMt) {
            HRESULT hrSMT = s->stillGrabber->SetMediaType(stillMt);
            dbg_printf("Fallback stillGrabber->SetMediaType(first STILL MT) => %s\n",
                HResultToString(hrSMT).c_str());
            if (stillMt->cbFormat && stillMt->pbFormat) CoTaskMemFree(stillMt->pbFormat);
            if (stillMt->pUnk) stillMt->pUnk->Release();
            CoTaskMemFree(stillMt);
        }
        emt->Release();
    }

    hrStill = s->cap->RenderStream(&PIN_CATEGORY_STILL, &MEDIATYPE_Video,
        s->capFilter, s->stillGrabberFilter, s->stillNullRenderer);
    dbg_printf("Fallback RenderStream(STILL, Video) => %s\n", HResultToString(hrStill).c_str());

    if (FAILED(hrStill)) {
        hrStill = s->cap->RenderStream(&PIN_CATEGORY_STILL, nullptr,
            s->capFilter, s->stillGrabberFilter, s->stillNullRenderer);
        dbg_printf("Fallback RenderStream(STILL, Any) => %s\n", HResultToString(hrStill).c_str());
    }

    // Manual fallback: explicit direct connect stillOut -> stillGrabber -> stillNull.
    if (FAILED(hrStill)) {
        auto find_first_pin = [](IBaseFilter* f, PIN_DIRECTION dir, IPin** out) -> HRESULT {
            *out = nullptr;
            IEnumPins* en = nullptr;
            HRESULT hrP = f->EnumPins(&en);
            if (FAILED(hrP) || !en) return FAILED(hrP) ? hrP : E_FAIL;
            IPin* p = nullptr;
            ULONG got = 0;
            while (en->Next(1, &p, &got) == S_OK) {
                PIN_DIRECTION d{};
                if (SUCCEEDED(p->QueryDirection(&d)) && d == dir) {
                    *out = p;
                    en->Release();
                    return S_OK;
                }
                p->Release();
            }
            en->Release();
            return HRESULT_FROM_WIN32(ERROR_NOT_FOUND);
        };

        IPin* grabIn = nullptr;
        IPin* grabOut = nullptr;
        IPin* nullIn = nullptr;
        HRESULT hrA = find_first_pin(s->stillGrabberFilter, PINDIR_INPUT, &grabIn);
        HRESULT hrB = find_first_pin(s->stillGrabberFilter, PINDIR_OUTPUT, &grabOut);
        HRESULT hrC = find_first_pin(s->stillNullRenderer, PINDIR_INPUT, &nullIn);

        dbg_printf("Fallback manual pin lookup: grabIn=%s grabOut=%s nullIn=%s\n",
            HResultToString(hrA).c_str(), HResultToString(hrB).c_str(), HResultToString(hrC).c_str());

        if (SUCCEEDED(hrA) && SUCCEEDED(hrB) && SUCCEEDED(hrC)) {
            HRESULT hr1 = s->graph->ConnectDirect(stillOut, grabIn, nullptr);
            dbg_printf("Fallback ConnectDirect(stillOut->grabIn) => %s\n", HResultToString(hr1).c_str());
            HRESULT hr2 = SUCCEEDED(hr1)
                ? s->graph->ConnectDirect(grabOut, nullIn, nullptr)
                : E_FAIL;
            dbg_printf("Fallback ConnectDirect(grabOut->nullIn) => %s\n", HResultToString(hr2).c_str());
            hrStill = (SUCCEEDED(hr1) && SUCCEEDED(hr2)) ? S_OK : FAILED(hr1) ? hr1 : hr2;
        }

        SAFE_RELEASE(grabIn);
        SAFE_RELEASE(grabOut);
        SAFE_RELEASE(nullIn);
    }

    SAFE_RELEASE(stillOut);

    if (FAILED(hrStill)) {
        cleanup_still_fallback_branch(s);
        return hrStill;
    }

    s->stillCbObj = new StillButtonCB(s);
    hrStill = s->stillGrabber->SetCallback(s->stillCbObj, 0);
    dbg_printf("Fallback stillGrabber->SetCallback => %s\n", HResultToString(hrStill).c_str());
    if (FAILED(hrStill)) {
        cleanup_still_fallback_branch(s);
        return hrStill;
    }

    return S_OK;
}

// ---- Build capture graph: BGRA compatibility output or direct native frames ----
static HRESULT build_capture_graph(
    DsSession* s,
    const DsDevice& dev,
    uint32_t streamCapsIndex,
    int outputMode)
{
    constexpr int kMaxStreamCapsBytes = 1024 * 1024;

    HRESULT hr;

    hr = CoCreateInstance(CLSID_FilterGraph, nullptr, CLSCTX_INPROC_SERVER, IID_IGraphBuilder, (void**)&s->graph);
    if (FAILED(hr)) return hr;

    hr = CoCreateInstance(CLSID_CaptureGraphBuilder2, nullptr, CLSCTX_INPROC_SERVER, IID_ICaptureGraphBuilder2, (void**)&s->cap);
    if (FAILED(hr)) return hr;

    hr = s->cap->SetFiltergraph(s->graph);
    if (FAILED(hr)) return hr;

    IMoniker* mk = nullptr;
    hr = bind_moniker_by_display_name(dev.monikerDisplayNameW, &mk);
    if (FAILED(hr)) return hr;

    hr = mk->BindToObject(nullptr, nullptr, IID_IBaseFilter, (void**)&s->capFilter);
    mk->Release();
    if (FAILED(hr)) return hr;

    hr = s->graph->AddFilter(s->capFilter, L"Capture");
    if (FAILED(hr)) return hr;

    DumpFilterPins(s->capFilter, "AfterAddFilter");

    HRESULT hrProcAmp = s->capFilter->QueryInterface(IID_IAMVideoProcAmp, (void**)&s->videoProcAmp);
    dbg_printf("QI(IAMVideoProcAmp) => %s\n", HResultToString(hrProcAmp).c_str());

    HRESULT hrCameraCtrl = s->capFilter->QueryInterface(IID_IAMCameraControl, (void**)&s->cameraControl);
    dbg_printf("QI(IAMCameraControl) => %s\n", HResultToString(hrCameraCtrl).c_str());

    // -----------------------------
    // IAMVideoControl trigger setup
    // -----------------------------
    {
        HRESULT hrVC = s->capFilter->QueryInterface(IID_IAMVideoControl, (void**)&s->videoCtrl);
        dbg_printf("QI(IAMVideoControl) => %s\n", HResultToString(hrVC).c_str());

        if (SUCCEEDED(hrVC) && s->videoCtrl) {
            IPin* stillOut = nullptr;
            HRESULT hrStillPin = FindPinByCategory(s->capFilter, PIN_CATEGORY_STILL, PINDIR_OUTPUT, &stillOut);
            dbg_printf("FindPinByCategory(STILL for IAMVideoControl) => %s\n", HResultToString(hrStillPin).c_str());

            if (SUCCEEDED(hrStillPin) && stillOut) {
                s->stillPinVC = stillOut; // keep ref

                long caps = 0;
                HRESULT hrCaps = s->videoCtrl->GetCaps(s->stillPinVC, &caps);
                dbg_printf("IAMVideoControl::GetCaps => %s caps=0x%08lx\n",
                    HResultToString(hrCaps).c_str(), caps);

                s->vcCaps = caps;
                s->vcHasTrigger = SUCCEEDED(hrCaps) && ((caps & VideoControlFlag_Trigger) != 0);

                long mode = 0;
                HRESULT hrMode = s->videoCtrl->GetMode(s->stillPinVC, &mode);
                dbg_printf("IAMVideoControl::GetMode => %s mode=0x%08lx\n",
                    HResultToString(hrMode).c_str(), mode);

                bool armAttempted = false;
                bool armSucceeded = false;

                if (SUCCEEDED(hrMode))
                    s->lastVcMode = mode;

                // Some UVC drivers latch Trigger high until user-mode clears it.
                // Arm by enabling external trigger (if supported) and clearing Trigger.
                if (SUCCEEDED(hrMode) && s->vcHasTrigger) {
                    long armMode = mode;
                    if ((caps & VideoControlFlag_ExternalTriggerEnable) != 0)
                        armMode |= VideoControlFlag_ExternalTriggerEnable;
                    armMode &= ~VideoControlFlag_Trigger;

                    if (armMode != mode) {
                        armAttempted = true;
                        HRESULT hrArm = s->videoCtrl->SetMode(s->stillPinVC, armMode);
                        dbg_printf("IAMVideoControl::SetMode(arm/clear trigger) => %s mode=0x%08lx\n",
                            HResultToString(hrArm).c_str(), armMode);
                        if (SUCCEEDED(hrArm)) {
                            armSucceeded = true;
                            long verifyMode = 0;
                            HRESULT hrVerify = s->videoCtrl->GetMode(s->stillPinVC, &verifyMode);
                            dbg_printf("IAMVideoControl::GetMode(after arm) => %s mode=0x%08lx\n",
                                HResultToString(hrVerify).c_str(), verifyMode);
                            if (SUCCEEDED(hrVerify)) {
                                s->lastVcMode = verifyMode;
                            }
                            else {
                                s->lastVcMode = armMode;
                            }
                        }
                    }
                }

                // If we cannot clear/arm trigger, fall back to STILL-sample callback path.
                if (s->vcHasTrigger && armAttempted && !armSucceeded) {
                    s->useStillFallback = true;
                }

                dbg_printf("IAMVideoControl trigger support: %s\n",
                    s->vcHasTrigger ? "YES" : "NO");
                dbg_printf("Trigger fallback via STILL callback: %s\n",
                    s->useStillFallback ? "YES" : "NO");
            }
            else {
                SAFE_RELEASE(s->videoCtrl);
            }
        }
    }

    // -----------------------------
    // Set device format (native)
    // -----------------------------
    IAMStreamConfig* cfg = nullptr;
    hr = s->cap->FindInterface(&PIN_CATEGORY_CAPTURE, &MEDIATYPE_Video, s->capFilter, IID_IAMStreamConfig, (void**)&cfg);
    if (FAILED(hr))
        hr = s->cap->FindInterface(&PIN_CATEGORY_PREVIEW, &MEDIATYPE_Video, s->capFilter, IID_IAMStreamConfig, (void**)&cfg);

    if (FAILED(hr) || !cfg) return E_FAIL;

    int count = 0, size = 0;
    HRESULT hrCaps = cfg->GetNumberOfCapabilities(&count, &size);
    if (FAILED(hrCaps) ||
        count <= 0 ||
        size < (int)sizeof(VIDEO_STREAM_CONFIG_CAPS) ||
        size > kMaxStreamCapsBytes) {
        SAFE_RELEASE(cfg);
        return E_FAIL;
    }

    if ((int)streamCapsIndex >= count) { SAFE_RELEASE(cfg); return E_FAIL; }

    AM_MEDIA_TYPE* mt = nullptr;
    std::vector<uint8_t> capsBuf((size_t)size);

    hr = cfg->GetStreamCaps((int)streamCapsIndex, &mt, capsBuf.data());
    if (FAILED(hr) || !mt) { SAFE_RELEASE(cfg); return E_FAIL; }

    if (outputMode != CDS_OUTPUT_BGRA && outputMode != CDS_OUTPUT_NATIVE) {
        free_am_media_type(mt);
        SAFE_RELEASE(cfg);
        return E_INVALIDARG;
    }

    const GUID selectedSubtype = mt->subtype;
    if (outputMode == CDS_OUTPUT_NATIVE &&
        selectedSubtype != MEDIASUBTYPE_NV12 &&
        selectedSubtype != MEDIASUBTYPE_YUY2 &&
        selectedSubtype != MEDIASUBTYPE_MJPG) {
        free_am_media_type(mt);
        SAFE_RELEASE(cfg);
        return HRESULT_FROM_WIN32(ERROR_NOT_SUPPORTED);
    }
    s->outputMode = outputMode;

    // GetStreamCaps can return a conservative default AvgTimePerFrame even
    // when the capability advertises a faster valid interval. Always request
    // the fastest rate for the explicitly selected capability.
    auto selectedCaps = (VIDEO_STREAM_CONFIG_CAPS*)capsBuf.data();
    if (selectedCaps->MinFrameInterval > 0) {
        if (set_media_type_frame_interval(mt, selectedCaps->MinFrameInterval)) {
            dbg_printf(
                "Requesting fastest frame interval: %lld (%.3f fps)\n",
                (long long)selectedCaps->MinFrameInterval,
                10000000.0 / (double)selectedCaps->MinFrameInterval);
        }
    }

    hr = cfg->SetFormat(mt);
    if (FAILED(hr)) {
        free_am_media_type(mt);
        SAFE_RELEASE(cfg);
        return hr;
    }

    if (!try_get_media_type_dimensions(mt, s->width, s->height)) {
        free_am_media_type(mt);
        SAFE_RELEASE(cfg);
        return E_FAIL;
    }

    free_am_media_type(mt);
    SAFE_RELEASE(cfg);

    if (!s->width || !s->height) return E_FAIL;

    // -----------------------------
    // SampleGrabber (BGRA compatibility output or selected native subtype)
    // -----------------------------
    hr = CoCreateInstance(__uuidof(CLSID_SampleGrabber), nullptr, CLSCTX_INPROC_SERVER,
        IID_IBaseFilter, (void**)&s->grabberFilter);
    if (FAILED(hr)) return hr;

    hr = s->graph->AddFilter(s->grabberFilter, L"FrameGrabber");
    if (FAILED(hr)) return hr;
    hr = s->grabberFilter->QueryInterface(__uuidof(ISampleGrabber), (void**)&s->grabber);
    if (FAILED(hr) || !s->grabber) return FAILED(hr) ? hr : E_FAIL;

    size_t rowBytes = 0;
    size_t frameBytes = 0;
    if (outputMode == CDS_OUTPUT_BGRA) {
        VIDEOINFOHEADER vih{};
        vih.bmiHeader.biSize = sizeof(BITMAPINFOHEADER);
        vih.bmiHeader.biWidth = (LONG)s->width;
        vih.bmiHeader.biHeight = -(LONG)s->height;
        vih.bmiHeader.biPlanes = 1;
        vih.bmiHeader.biBitCount = 32;
        vih.bmiHeader.biCompression = BI_RGB;
        if (!calc_frame_layout_bytes(s->width, s->height, rowBytes, frameBytes)) return E_FAIL;
        if (rowBytes > (size_t)(std::numeric_limits<int32_t>::max)()) return E_FAIL;
        if (frameBytes > (std::numeric_limits<DWORD>::max)()) return E_FAIL;
        vih.bmiHeader.biSizeImage = (DWORD)frameBytes;

        AM_MEDIA_TYPE rgb{};
        rgb.majortype = MEDIATYPE_Video;
        rgb.subtype = MEDIASUBTYPE_RGB32;
        rgb.formattype = FORMAT_VideoInfo;
        rgb.cbFormat = sizeof(VIDEOINFOHEADER);
        rgb.pbFormat = (BYTE*)CoTaskMemAlloc(sizeof(VIDEOINFOHEADER));
        if (!rgb.pbFormat) return E_OUTOFMEMORY;
        memcpy(rgb.pbFormat, &vih, sizeof(VIDEOINFOHEADER));

        hr = s->grabber->SetMediaType(&rgb);
        CoTaskMemFree(rgb.pbFormat);
    }
    else {
        AM_MEDIA_TYPE native{};
        native.majortype = MEDIATYPE_Video;
        native.subtype = selectedSubtype;
        native.formattype = GUID_NULL;
        hr = s->grabber->SetMediaType(&native);
    }
    if (FAILED(hr)) return hr;

    hr = s->grabber->SetBufferSamples(FALSE);
    if (FAILED(hr)) return hr;

    hr = CoCreateInstance(__uuidof(CLSID_NullRenderer), nullptr, CLSCTX_INPROC_SERVER,
        IID_IBaseFilter, (void**)&s->nullRenderer);
    if (FAILED(hr)) return hr;

    hr = s->graph->AddFilter(s->nullRenderer, L"NullRenderer");
    if (FAILED(hr)) return hr;

    hr = s->cap->RenderStream(&PIN_CATEGORY_CAPTURE, &MEDIATYPE_Video,
        s->capFilter, s->grabberFilter, s->nullRenderer);

    if (FAILED(hr))
        hr = s->cap->RenderStream(&PIN_CATEGORY_PREVIEW, &MEDIATYPE_Video,
            s->capFilter, s->grabberFilter, s->nullRenderer);

    if (FAILED(hr)) return hr;

    // Detect the actual connected subtype/layout. Positive biHeight only means
    // bottom-up for RGB; YUV formats are copied in their native top-down layout.
    s->bottomUp = false;
    {
        AM_MEDIA_TYPE connected{};
        if (SUCCEEDED(s->grabber->GetConnectedMediaType(&connected))) {
            uint32_t connectedWidth = s->width;
            uint32_t connectedHeight = s->height;
            (void)try_get_media_type_dimensions(&connected, connectedWidth, connectedHeight);
            s->width = connectedWidth;
            s->height = connectedHeight;

            REFERENCE_TIME connectedInterval = media_type_frame_interval(&connected);
            if (connectedInterval > 0) {
                dbg_printf(
                    "Connected frame interval: %lld (%.3f fps)\n",
                    (long long)connectedInterval,
                    10000000.0 / (double)connectedInterval);
            }

            if (outputMode == CDS_OUTPUT_NATIVE && connected.subtype != selectedSubtype) {
                if (connected.cbFormat && connected.pbFormat) CoTaskMemFree(connected.pbFormat);
                if (connected.pUnk) connected.pUnk->Release();
                return HRESULT_FROM_WIN32(ERROR_NOT_SUPPORTED);
            }

            if (!calc_connected_frame_layout(
                    &connected,
                    s->width,
                    s->height,
                    s->frameLayout)) {
                if (connected.cbFormat && connected.pbFormat) CoTaskMemFree(connected.pbFormat);
                if (connected.pUnk) connected.pUnk->Release();
                return HRESULT_FROM_WIN32(ERROR_NOT_SUPPORTED);
            }
            frameBytes = s->frameLayout.dataBytes;
            rowBytes = s->frameLayout.planeStrides[0];

            if (s->frameLayout.pixelFormat == CDS_PIXEL_FORMAT_BGRA) {
                const BITMAPINFOHEADER* bmi = media_type_bitmap_header(&connected);
                if (bmi) {
                    s->bottomUp = bmi->biHeight > 0;
                    dbg_printf("RGB orientation: biHeight=%ld -> bottomUp=%s\n",
                        (long)bmi->biHeight, s->bottomUp ? "YES" : "NO");
                }
            }
            dbg_printf(
                "Connected output: %s bytes=%zu stride0=%zu offset1=%zu stride1=%zu\n",
                PixelFormatName(s->frameLayout.pixelFormat),
                s->frameLayout.dataBytes,
                s->frameLayout.planeStrides[0],
                s->frameLayout.planeOffsets[1],
                s->frameLayout.planeStrides[1]);

            if (connected.cbFormat && connected.pbFormat) CoTaskMemFree(connected.pbFormat);
            if (connected.pUnk) connected.pUnk->Release();
        }
        else {
            return E_FAIL;
        }
    }

    s->frameCbObj = new FrameGrabberCB(s);
    if (!s->frameCbObj) return E_OUTOFMEMORY;
    hr = s->grabber->SetCallback(s->frameCbObj, 1);
    if (FAILED(hr)) return hr;

    if (s->useStillFallback) {
        HRESULT hrStill = build_still_fallback_button_branch(s);
        dbg_printf("Fallback build STILL branch => %s\n", HResultToString(hrStill).c_str());
        if (FAILED(hrStill)) {
            dbg_printf("Fallback STILL path unavailable; button events may be unavailable.\n");
            s->useStillFallback = false;
        }
    }

    hr = s->graph->QueryInterface(IID_IMediaControl, (void**)&s->mc);
    if (FAILED(hr) || !s->mc) return FAILED(hr) ? hr : E_FAIL;
    s->graph->QueryInterface(IID_IMediaEvent, (void**)&s->me);

    s->lastFrame.resize(frameBytes);

    return S_OK;
}

static void session_thread_main(
    DsSession* s,
    DsDevice devCopy,
    uint32_t streamCapsIndex,
    int outputMode) {
    CoInitializeEx(nullptr, COINIT_APARTMENTTHREADED | COINIT_DISABLE_OLE1DDE);

    auto signal_start = [&](cds_result_t r) {
        {
            std::lock_guard<std::mutex> lk(s->startMutex);
            if (!s->startCompleted) {
                s->startResult = r;
                s->startCompleted = true;
            }
        }
        s->startCv.notify_all();
    };

    HRESULT hr = build_capture_graph(s, devCopy, streamCapsIndex, outputMode);
    if (SUCCEEDED(hr)) {
        if (!s->mc) {
            signal_start(CDS_ERR_OPENING_DEVICE);
            dbg_printf("cds: IMediaControl missing after graph build\n");
            s->release_graph_thread_only();
            CoUninitialize();
            return;
        }

        hr = s->mc->Run();
        if (FAILED(hr)) {
            signal_start(CDS_ERR_OPENING_DEVICE);
            dbg_printf("cds: Run failed: %s\n", HResultToString(hr).c_str());
        }
        else {
            signal_start(CDS_OK);
            while (!s->stopRequested.load()) {
                process_pending_control_request(s);

                if (!s->useStillFallback && s->vcHasTrigger && s->videoCtrl && s->stillPinVC) {
                    long mode = 0;
                    HRESULT hrMode = s->videoCtrl->GetMode(s->stillPinVC, &mode);
                    if (SUCCEEDED(hrMode)) {
                        bool was = (s->lastVcMode & VideoControlFlag_Trigger) != 0;
                        bool now = (mode & VideoControlFlag_Trigger) != 0;

                        if (!was && now) {
                            s->lastButtonTs100ns.store(now_ts100ns_utc());
                            s->buttonEdge.store(true);
                            dbg_printf("[UVC TRIGGER] rising edge\n");

                            // Re-arm for devices that latch Trigger until cleared.
                            long clearMode = mode & ~VideoControlFlag_Trigger;
                            HRESULT hrClear = s->videoCtrl->SetMode(s->stillPinVC, clearMode);
                            dbg_printf("IAMVideoControl::SetMode(clear trigger) => %s mode=0x%08lx\n",
                                HResultToString(hrClear).c_str(), clearMode);
                            if (SUCCEEDED(hrClear)) {
                                long verifyMode = 0;
                                HRESULT hrVerify = s->videoCtrl->GetMode(s->stillPinVC, &verifyMode);
                                if (SUCCEEDED(hrVerify)) {
                                    s->lastVcMode = verifyMode;
                                    continue;
                                }
                                s->lastVcMode = clearMode;
                                continue;
                            }
                        }

                        s->lastVcMode = mode;
                    }
                }

                Sleep(5);
            }

            // SAFE STOP (same thread)
            s->mc->Stop();

            // Drain events (optional)
            if (s->me) {
                long ev = 0;
                LONG_PTR p1 = 0, p2 = 0;
                while (s->me->GetEvent(&ev, &p1, &p2, 0) == S_OK) {
                    s->me->FreeEventParams(ev, p1, p2);
                }
            }
        }
    }
    else {
        const cds_result_t buildResult =
            hr == HRESULT_FROM_WIN32(ERROR_NOT_SUPPORTED)
                ? CDS_ERR_FORMAT_NOT_SUPPORTED
                : (hr == E_INVALIDARG ? CDS_ERR_INVALID_ARGUMENT : CDS_ERR_OPENING_DEVICE);
        signal_start(buildResult);
        dbg_printf("cds: build graph failed: %s\n", HResultToString(hr).c_str());
    }

    // Ensure waiter is always released even on unexpected paths.
    signal_start(CDS_ERR_OPENING_DEVICE);
    fail_pending_control_request(s, HRESULT_FROM_WIN32(ERROR_OPERATION_ABORTED));

    // FULL TEARDOWN (same thread)
    s->release_graph_thread_only();
    CoUninitialize();
}

static cds_result_t export_get_control_range(
    DsControlApi api,
    int32_t device_index,
    int32_t property,
    int32_t* min_value,
    int32_t* max_value,
    int32_t* step,
    int32_t* default_value,
    int32_t* caps_flags)
{
    if (!min_value || !max_value || !step || !default_value || !caps_flags) {
        return CDS_ERR_INVALID_ARGUMENT;
    }

    DsControlPayload payload{};
    cds_result_t rc = execute_device_control_request(
        device_index,
        api,
        DsControlRequestType::GetRange,
        (long)property,
        &payload);
    if (rc != CDS_OK) return rc;

    *min_value = (int32_t)payload.minValue;
    *max_value = (int32_t)payload.maxValue;
    *step = (int32_t)payload.step;
    *default_value = (int32_t)payload.defaultValue;
    *caps_flags = (int32_t)payload.capabilities;
    return CDS_OK;
}

static cds_result_t export_get_control(
    DsControlApi api,
    int32_t device_index,
    int32_t property,
    int32_t* value,
    int32_t* flags)
{
    if (!value || !flags) return CDS_ERR_INVALID_ARGUMENT;

    DsControlPayload payload{};
    cds_result_t rc = execute_device_control_request(
        device_index,
        api,
        DsControlRequestType::Get,
        (long)property,
        &payload);
    if (rc != CDS_OK) return rc;

    *value = (int32_t)payload.value;
    *flags = (int32_t)payload.flags;
    return CDS_OK;
}

static cds_result_t export_set_control(
    DsControlApi api,
    int32_t device_index,
    int32_t property,
    int32_t value,
    int32_t flags)
{
    DsControlPayload payload{};
    payload.value = (long)value;
    payload.flags = (long)flags;

    return execute_device_control_request(
        device_index,
        api,
        DsControlRequestType::Set,
        (long)property,
        &payload);
}

// =============================================================================
// ============================== C API Exports ===============================
// =============================================================================

extern "C" {
    // -------------------- cds_* exports --------------------

    SP_API cds_result_t SP_CALL cds_initialize(void) {
        std::lock_guard<std::mutex> lk(g_dsMutex);

        HRESULT hr = CoInitializeEx(nullptr, COINIT_MULTITHREADED);
        bool didInit = SUCCEEDED(hr);
        if (FAILED(hr) && hr != RPC_E_CHANGED_MODE) {
            return CDS_ERR_UNKNOWN;
        }

        hr = enumerate_devices_and_formats();
        if (FAILED(hr)) {
            if (didInit) CoUninitialize();
            return CDS_ERR_UNKNOWN;
        }

        if (didInit) CoUninitialize();
        g_dsInitialized = true;
        ++g_dsGeneration;
        return CDS_OK;
    }

    SP_API void SP_CALL cds_shutdown_capture_api(void) {
        // Stop all sessions first (outside lock join)
        std::vector<uint32_t> toStop;
        {
            std::lock_guard<std::mutex> lk(g_dsMutex);
            g_dsInitialized = false;
            ++g_dsGeneration;
            for (auto& kv : g_dsSessions) toStop.push_back(kv.first);
        }
        for (auto idx : toStop) {
            cds_stop_capture(idx);
        }

        std::lock_guard<std::mutex> lk(g_dsMutex);
        g_dsSessions.clear();
        g_dsDevices.clear();
        g_dsInitialized = false;
    }

    SP_API void SP_CALL cds_set_log_enabled(int32_t enabled) {
        g_logOverride.store(enabled ? 1 : 0, std::memory_order_relaxed);
    }

    SP_API int32_t SP_CALL cds_devices_count(void) {
        std::lock_guard<std::mutex> lk(g_dsMutex);
        if (!g_dsInitialized) return 0;
        return (int32_t)g_dsDevices.size();
    }

    SP_API size_t SP_CALL cds_device_name(int32_t device_index, char* buf, size_t buf_len) {
        std::lock_guard<std::mutex> lk(g_dsMutex);
        if (!g_dsInitialized) return 0;
        if (device_index < 0 || (size_t)device_index >= g_dsDevices.size()) return 0;
        return copy_str(g_dsDevices[(size_t)device_index].nameUtf8, buf, buf_len);
    }

    SP_API size_t SP_CALL cds_device_unique_id(int32_t device_index, char* buf, size_t buf_len) {
        std::lock_guard<std::mutex> lk(g_dsMutex);
        if (!g_dsInitialized) return 0;
        if (device_index < 0 || (size_t)device_index >= g_dsDevices.size()) return 0;
        return copy_str(g_dsDevices[(size_t)device_index].devicePathUtf8, buf, buf_len);
    }

    SP_API size_t SP_CALL cds_device_model_id(int32_t device_index, char* buf, size_t buf_len) {
        std::lock_guard<std::mutex> lk(g_dsMutex);
        if (!g_dsInitialized) return 0;
        if (device_index < 0 || (size_t)device_index >= g_dsDevices.size()) return 0;
        return copy_str(g_dsDevices[(size_t)device_index].modelIdUtf8, buf, buf_len);
    }

    SP_API int32_t SP_CALL cds_device_vid(int32_t device_index) {
        std::lock_guard<std::mutex> lk(g_dsMutex);
        if (!g_dsInitialized) return 0;
        if (device_index < 0 || (size_t)device_index >= g_dsDevices.size()) return 0;
        return g_dsDevices[(size_t)device_index].vid;
    }

    SP_API int32_t SP_CALL cds_device_pid(int32_t device_index) {
        std::lock_guard<std::mutex> lk(g_dsMutex);
        if (!g_dsInitialized) return 0;
        if (device_index < 0 || (size_t)device_index >= g_dsDevices.size()) return 0;
        return g_dsDevices[(size_t)device_index].pid;
    }

    SP_API int32_t SP_CALL cds_device_formats_count(int32_t device_index) {
        std::lock_guard<std::mutex> lk(g_dsMutex);
        if (!g_dsInitialized) return CDS_ERR_NOT_INITIALIZED;
        if (device_index < 0 || (size_t)device_index >= g_dsDevices.size()) return CDS_ERR_DEVICE_NOT_FOUND;
        return (int32_t)g_dsDevices[(size_t)device_index].formats.size();
    }

    SP_API uint32_t SP_CALL cds_device_format_width(int32_t device_index, int32_t format_index) {
        std::lock_guard<std::mutex> lk(g_dsMutex);
        if (!g_dsInitialized) return 0;
        if (device_index < 0 || (size_t)device_index >= g_dsDevices.size()) return 0;
        auto& v = g_dsDevices[(size_t)device_index].formats;
        if (format_index < 0 || (size_t)format_index >= v.size()) return 0;
        return v[(size_t)format_index].width;
    }

    SP_API uint32_t SP_CALL cds_device_format_height(int32_t device_index, int32_t format_index) {
        std::lock_guard<std::mutex> lk(g_dsMutex);
        if (!g_dsInitialized) return 0;
        if (device_index < 0 || (size_t)device_index >= g_dsDevices.size()) return 0;
        auto& v = g_dsDevices[(size_t)device_index].formats;
        if (format_index < 0 || (size_t)format_index >= v.size()) return 0;
        return v[(size_t)format_index].height;
    }

    SP_API uint32_t SP_CALL cds_device_format_frame_rate(int32_t device_index, int32_t format_index) {
        std::lock_guard<std::mutex> lk(g_dsMutex);
        if (!g_dsInitialized) return 0;
        if (device_index < 0 || (size_t)device_index >= g_dsDevices.size()) return 0;
        auto& v = g_dsDevices[(size_t)device_index].formats;
        if (format_index < 0 || (size_t)format_index >= v.size()) return 0;
        return v[(size_t)format_index].maxFps;
    }

    SP_API size_t SP_CALL cds_device_format_type(int32_t device_index, int32_t format_index, char* buf, size_t buf_len) {
        std::lock_guard<std::mutex> lk(g_dsMutex);
        if (!g_dsInitialized) return 0;
        if (device_index < 0 || (size_t)device_index >= g_dsDevices.size()) return 0;
        auto& v = g_dsDevices[(size_t)device_index].formats;
        if (format_index < 0 || (size_t)format_index >= v.size()) return 0;

        const char* name = SubTypeName(v[(size_t)format_index].subtype);
        if (name) return copy_str(name, buf, buf_len);
        return copy_str(GuidToStr(v[(size_t)format_index].subtype), buf, buf_len);
    }

    SP_API cds_result_t SP_CALL cds_start_capture(uint32_t device_index, uint32_t width, uint32_t height) {
        return cds_start_capture_with_output(device_index, width, height, CDS_OUTPUT_BGRA);
    }

    SP_API cds_result_t SP_CALL cds_start_capture_with_output(
        uint32_t device_index,
        uint32_t width,
        uint32_t height,
        int32_t output_mode)
    {
        if (output_mode != CDS_OUTPUT_BGRA && output_mode != CDS_OUTPUT_NATIVE) {
            return CDS_ERR_INVALID_ARGUMENT;
        }

        std::vector<uint32_t> rankedFormatIndices;
        {
            std::lock_guard<std::mutex> lk(g_dsMutex);
            if (!g_dsInitialized) return CDS_ERR_NOT_INITIALIZED;
            if (device_index >= g_dsDevices.size()) return CDS_ERR_DEVICE_NOT_FOUND;
            if (g_dsSessions.count(device_index)) return CDS_ERR_ALREADY_STARTED;

            auto& fmts = g_dsDevices[device_index].formats;

            for (int i = 0; i < (int)fmts.size(); ++i) {
                if (fmts[i].width != width || fmts[i].height != height) continue;
                rankedFormatIndices.push_back((uint32_t)i);
            }

            if (rankedFormatIndices.empty()) return CDS_ERR_FORMAT_NOT_FOUND;
            std::stable_sort(
                rankedFormatIndices.begin(),
                rankedFormatIndices.end(),
                [&](uint32_t a, uint32_t b) {
                    return cds_format_selection::IsBetter(
                        cds_format_selection::Score(
                            fmts[(size_t)a].maxFps,
                            SubTypeProcessingRank(fmts[(size_t)a].subtype),
                            a),
                        cds_format_selection::Score(
                            fmts[(size_t)b].maxFps,
                            SubTypeProcessingRank(fmts[(size_t)b].subtype),
                            b));
                });

            for (size_t rank = 0; rank < rankedFormatIndices.size(); ++rank) {
                uint32_t index = rankedFormatIndices[rank];
                const char* subtype = SubTypeName(fmts[(size_t)index].subtype);
                dbg_printf(
                    "Resolution-only candidate #%zu: %ux%u index=%u fps=%u subtype=%s processing-rank=%d\n",
                    rank + 1,
                    width,
                    height,
                    index,
                    fmts[(size_t)index].maxFps,
                    subtype ? subtype : "OTHER",
                    SubTypeProcessingRank(fmts[(size_t)index].subtype));
            }
        }

        cds_result_t lastResult = CDS_ERR_OPENING_DEVICE;
        for (uint32_t formatIndex : rankedFormatIndices) {
            lastResult = cds_start_capture_with_format_output(
                device_index,
                formatIndex,
                output_mode);
            if (lastResult == CDS_OK) {
                dbg_printf(
                    "Resolution-only selection succeeded with format index=%u\n",
                    formatIndex);
                return CDS_OK;
            }
            if (lastResult != CDS_ERR_OPENING_DEVICE &&
                lastResult != CDS_ERR_FORMAT_NOT_SUPPORTED) {
                return lastResult;
            }
            dbg_printf(
                "Resolution-only format index=%u could not build/run; trying next candidate\n",
                formatIndex);
        }
        return lastResult;
    }

    SP_API cds_result_t SP_CALL cds_start_capture_with_format(uint32_t device_index, uint32_t format_index) {
        return cds_start_capture_with_format_output(
            device_index,
            format_index,
            CDS_OUTPUT_BGRA);
    }

    SP_API cds_result_t SP_CALL cds_start_capture_with_format_output(
        uint32_t device_index,
        uint32_t format_index,
        int32_t output_mode)
    {
        if (output_mode != CDS_OUTPUT_BGRA && output_mode != CDS_OUTPUT_NATIVE) {
            return CDS_ERR_INVALID_ARGUMENT;
        }

        DsDevice devCopy;
        uint32_t streamCapsIndex = 0;
        uint64_t generationSnapshot = 0;

        {
            std::lock_guard<std::mutex> lk(g_dsMutex);
            if (!g_dsInitialized) return CDS_ERR_NOT_INITIALIZED;
            if (device_index >= g_dsDevices.size()) return CDS_ERR_DEVICE_NOT_FOUND;
            if (g_dsSessions.count(device_index)) return CDS_ERR_ALREADY_STARTED;
            if (format_index >= g_dsDevices[device_index].formats.size()) return CDS_ERR_FORMAT_NOT_FOUND;

            generationSnapshot = g_dsGeneration;
            devCopy = g_dsDevices[device_index];
            if (output_mode == CDS_OUTPUT_NATIVE) {
                const GUID& subtype = g_dsDevices[device_index].formats[format_index].subtype;
                if (subtype != MEDIASUBTYPE_NV12 &&
                    subtype != MEDIASUBTYPE_YUY2 &&
                    subtype != MEDIASUBTYPE_MJPG) {
                    return CDS_ERR_FORMAT_NOT_SUPPORTED;
                }
            }
            streamCapsIndex = g_dsDevices[device_index].formats[format_index].streamCapsIndex;
        }

        DsSession* s = new(std::nothrow) DsSession();
        if (!s) return CDS_ERR_UNKNOWN;

        s->deviceIndex = device_index;
        s->stopRequested.store(false);
        try {
            s->worker = std::thread(
                session_thread_main,
                s,
                devCopy,
                streamCapsIndex,
                output_mode);
        }
        catch (...) {
            delete s;
            return CDS_ERR_UNKNOWN;
        }

        cds_result_t startRc = CDS_ERR_UNKNOWN;
        {
            std::unique_lock<std::mutex> lk(s->startMutex);
            s->startCv.wait(lk, [&]() { return s->startCompleted; });
            startRc = s->startResult;
        }

        if (startRc != CDS_OK) {
            s->stopRequested.store(true);
            if (s->worker.joinable()) s->worker.join();
            delete s;
            return startRc;
        }

        bool rejectedNotInitialized = false;
        bool rejectedAlreadyStarted = false;
        {
            std::lock_guard<std::mutex> lk(g_dsMutex);
            if (!g_dsInitialized || generationSnapshot != g_dsGeneration) {
                rejectedNotInitialized = true;
            }
            else if (g_dsSessions.count(device_index)) {
                rejectedAlreadyStarted = true;
            }
            else {
                g_dsSessions[device_index] = s;
            }
        }
        if (rejectedNotInitialized || rejectedAlreadyStarted) {
            s->stopRequested.store(true);
            if (s->worker.joinable()) s->worker.join();
            delete s;
            return rejectedNotInitialized ? CDS_ERR_NOT_INITIALIZED : CDS_ERR_ALREADY_STARTED;
        }

        return CDS_OK;
    }

    SP_API cds_result_t SP_CALL cds_stop_capture(uint32_t device_index) {
        DsSession* s = nullptr;
        {
            std::lock_guard<std::mutex> lk(g_dsMutex);
            auto it = g_dsSessions.find(device_index);
            if (it == g_dsSessions.end()) return CDS_ERR_NOT_STARTED;
            s = it->second;
            g_dsSessions.erase(it);
        }

        s->stopRequested.store(true);
        if (s->worker.joinable()) s->worker.join();
        delete s;
        return CDS_OK;
    }

    SP_API int32_t SP_CALL cds_has_first_frame(uint32_t device_index) {
        std::lock_guard<std::mutex> lk(g_dsMutex);
        auto it = g_dsSessions.find(device_index);
        if (it == g_dsSessions.end()) return 0;
        return it->second->hasFrame.load() ? 1 : 0;
    }

    SP_API cds_result_t SP_CALL cds_grab_frame(uint32_t device_index, uint8_t* buffer, size_t available_bytes) {
        std::lock_guard<std::mutex> lk(g_dsMutex);
        auto it = g_dsSessions.find(device_index);
        if (it == g_dsSessions.end()) return CDS_ERR_NOT_STARTED;
        DsSession* s = it->second;

        if (!buffer) return CDS_ERR_BUF_NULL;

        bool callbackExclusive = false;
        {
            std::lock_guard<std::mutex> deliveryLock(s->deliveryMutex);
            callbackExclusive = s->frameCallback && s->frameCallbackExclusive;
        }

        std::unique_lock<std::mutex> lk2(s->frameMutex);
        size_t needed = s->frameLayout.dataBytes;
        if (needed == 0 || available_bytes < needed) return CDS_ERR_BUF_TOO_SMALL;

        if (callbackExclusive) {
            const uint64_t previousSequence = s->pullFrameSequence;
            s->pullFrameRequested.store(true, std::memory_order_release);
            const bool received = s->frameCv.wait_for(
                lk2,
                std::chrono::milliseconds(1000),
                [&]() {
                    return s->pullFrameSequence > previousSequence ||
                        s->stopRequested.load(std::memory_order_acquire);
                });
            if (!received || s->pullFrameSequence <= previousSequence) {
                return CDS_ERR_READ_FRAME;
            }
            // MJPEG samples are variable-sized. Re-read the size after the
            // requested sample arrives so we never silently truncate it.
            needed = s->frameLayout.dataBytes;
            if (needed == 0 || available_bytes < needed) {
                return CDS_ERR_BUF_TOO_SMALL;
            }
        }

        if (!s->pullFrameValid.load(std::memory_order_acquire) || s->lastFrame.size() < needed) {
            return CDS_ERR_READ_FRAME;
        }

        memcpy(buffer, s->lastFrame.data(), needed);
        return CDS_OK;
    }

    SP_API cds_result_t SP_CALL cds_set_frame_callback(
        uint32_t device_index,
        cds_frame_callback_t callback,
        void* user_data,
        int32_t exclusive)
    {
        std::lock_guard<std::mutex> sessionLock(g_dsMutex);
        auto it = g_dsSessions.find(device_index);
        if (it == g_dsSessions.end()) return CDS_ERR_NOT_STARTED;
        DsSession* s = it->second;

        std::unique_lock<std::mutex> deliveryLock(s->deliveryMutex);
        s->frameCallback = nullptr;
        s->frameCallbackUserData = nullptr;
        s->frameCallbackExclusive = false;
        s->deliveryCv.wait(deliveryLock, [&]() {
            return s->activeFrameCallbacks == 0;
        });

        if (callback) {
            s->frameCallbackUserData = user_data;
            s->frameCallbackExclusive = exclusive != 0;
            s->frameCallback = callback;
            if (s->frameCallbackExclusive) {
                s->pullFrameValid.store(false, std::memory_order_release);
                s->pullFrameRequested.store(false, std::memory_order_release);
            }
        }
        else {
            s->pullFrameValid.store(false, std::memory_order_release);
            s->pullFrameRequested.store(false, std::memory_order_release);
        }
        return CDS_OK;
    }

    SP_API int32_t SP_CALL cds_frame_width(uint32_t device_index) {
        std::lock_guard<std::mutex> lk(g_dsMutex);
        auto it = g_dsSessions.find(device_index);
        if (it == g_dsSessions.end()) return 0;
        return (int32_t)it->second->width;
    }

    SP_API int32_t SP_CALL cds_frame_height(uint32_t device_index) {
        std::lock_guard<std::mutex> lk(g_dsMutex);
        auto it = g_dsSessions.find(device_index);
        if (it == g_dsSessions.end()) return 0;
        return (int32_t)it->second->height;
    }

    SP_API int32_t SP_CALL cds_frame_bytes_per_row(uint32_t device_index) {
        std::lock_guard<std::mutex> lk(g_dsMutex);
        auto it = g_dsSessions.find(device_index);
        if (it == g_dsSessions.end()) return 0;
        std::lock_guard<std::mutex> lk2(it->second->frameMutex);
        const size_t stride = it->second->frameLayout.planeStrides[0];
        return stride <= (size_t)(std::numeric_limits<int32_t>::max)()
            ? (int32_t)stride
            : 0;
    }

    SP_API int32_t SP_CALL cds_frame_data_size(uint32_t device_index) {
        std::lock_guard<std::mutex> lk(g_dsMutex);
        auto it = g_dsSessions.find(device_index);
        if (it == g_dsSessions.end()) return 0;
        std::lock_guard<std::mutex> lk2(it->second->frameMutex);
        const size_t bytes = it->second->frameLayout.dataBytes;
        return bytes <= (size_t)(std::numeric_limits<int32_t>::max)()
            ? (int32_t)bytes
            : 0;
    }

    SP_API int32_t SP_CALL cds_frame_pixel_format(uint32_t device_index) {
        std::lock_guard<std::mutex> lk(g_dsMutex);
        auto it = g_dsSessions.find(device_index);
        if (it == g_dsSessions.end()) return CDS_PIXEL_FORMAT_UNKNOWN;
        std::lock_guard<std::mutex> lk2(it->second->frameMutex);
        return it->second->frameLayout.pixelFormat;
    }

    SP_API size_t SP_CALL cds_frame_pixel_format_name(
        uint32_t device_index,
        char* buf,
        size_t buf_len)
    {
        return copy_str(PixelFormatName(cds_frame_pixel_format(device_index)), buf, buf_len);
    }

    SP_API int32_t SP_CALL cds_frame_plane_count(uint32_t device_index) {
        std::lock_guard<std::mutex> lk(g_dsMutex);
        auto it = g_dsSessions.find(device_index);
        if (it == g_dsSessions.end()) return 0;
        std::lock_guard<std::mutex> lk2(it->second->frameMutex);
        return it->second->frameLayout.planeCount;
    }

    SP_API int32_t SP_CALL cds_frame_plane_offset(uint32_t device_index, int32_t plane_index) {
        std::lock_guard<std::mutex> lk(g_dsMutex);
        auto it = g_dsSessions.find(device_index);
        if (it == g_dsSessions.end()) return -1;
        std::lock_guard<std::mutex> lk2(it->second->frameMutex);
        if (plane_index < 0 || plane_index >= it->second->frameLayout.planeCount) return -1;
        const size_t offset = it->second->frameLayout.planeOffsets[plane_index];
        return offset <= (size_t)(std::numeric_limits<int32_t>::max)()
            ? (int32_t)offset
            : -1;
    }

    SP_API int32_t SP_CALL cds_frame_plane_bytes_per_row(
        uint32_t device_index,
        int32_t plane_index)
    {
        std::lock_guard<std::mutex> lk(g_dsMutex);
        auto it = g_dsSessions.find(device_index);
        if (it == g_dsSessions.end()) return 0;
        std::lock_guard<std::mutex> lk2(it->second->frameMutex);
        if (plane_index < 0 || plane_index >= it->second->frameLayout.planeCount) return 0;
        const size_t stride = it->second->frameLayout.planeStrides[plane_index];
        return stride <= (size_t)(std::numeric_limits<int32_t>::max)()
            ? (int32_t)stride
            : 0;
    }

    SP_API cds_result_t SP_CALL cds_get_video_proc_amp_range(
        int32_t device_index,
        int32_t property,
        int32_t* min_value,
        int32_t* max_value,
        int32_t* step,
        int32_t* default_value,
        int32_t* caps_flags)
    {
        return export_get_control_range(
            DsControlApi::VideoProcAmp,
            device_index,
            property,
            min_value,
            max_value,
            step,
            default_value,
            caps_flags);
    }

    SP_API cds_result_t SP_CALL cds_get_video_proc_amp(
        int32_t device_index,
        int32_t property,
        int32_t* value,
        int32_t* flags)
    {
        return export_get_control(
            DsControlApi::VideoProcAmp,
            device_index,
            property,
            value,
            flags);
    }

    SP_API cds_result_t SP_CALL cds_set_video_proc_amp(
        int32_t device_index,
        int32_t property,
        int32_t value,
        int32_t flags)
    {
        return export_set_control(
            DsControlApi::VideoProcAmp,
            device_index,
            property,
            value,
            flags);
    }

    SP_API cds_result_t SP_CALL cds_get_camera_control_range(
        int32_t device_index,
        int32_t property,
        int32_t* min_value,
        int32_t* max_value,
        int32_t* step,
        int32_t* default_value,
        int32_t* caps_flags)
    {
        return export_get_control_range(
            DsControlApi::CameraControl,
            device_index,
            property,
            min_value,
            max_value,
            step,
            default_value,
            caps_flags);
    }

    SP_API cds_result_t SP_CALL cds_get_camera_control(
        int32_t device_index,
        int32_t property,
        int32_t* value,
        int32_t* flags)
    {
        return export_get_control(
            DsControlApi::CameraControl,
            device_index,
            property,
            value,
            flags);
    }

    SP_API cds_result_t SP_CALL cds_set_camera_control(
        int32_t device_index,
        int32_t property,
        int32_t value,
        int32_t flags)
    {
        return export_set_control(
            DsControlApi::CameraControl,
            device_index,
            property,
            value,
            flags);
    }

    // Button while streaming: edge-trigger (1 once)
    SP_API int32_t SP_CALL cds_button_pressed(uint32_t device_index) {
        std::lock_guard<std::mutex> lk(g_dsMutex);
        auto it = g_dsSessions.find(device_index);
        if (it == g_dsSessions.end()) return 0;
        return it->second->buttonEdge.exchange(false) ? 1 : 0;
    }

    SP_API uint64_t SP_CALL cds_button_timestamp(uint32_t device_index) {
        std::lock_guard<std::mutex> lk(g_dsMutex);
        auto it = g_dsSessions.find(device_index);
        if (it == g_dsSessions.end()) return 0;
        return it->second->lastButtonTs100ns.load();
    }

} // extern "C"

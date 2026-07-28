#pragma once
#include <stdint.h>
#include <stddef.h>

#ifdef __cplusplus
extern "C" {
#endif

#ifdef _WIN32
#define SP_API __declspec(dllexport)
#define SP_CALL __stdcall
#else
#define SP_API
#define SP_CALL
#endif

	// ===================== DirectShow Capture API (cds_*) =====================

	typedef int32_t cds_result_t;

#define CDS_OK 0
#define CDS_ERR_DEVICE_NOT_FOUND -1
#define CDS_ERR_FORMAT_NOT_FOUND -2
#define CDS_ERR_OPENING_DEVICE   -3
#define CDS_ERR_ALREADY_STARTED  -4
#define CDS_ERR_NOT_STARTED      -5
#define CDS_ERR_NOT_INITIALIZED  -6
#define CDS_ERR_INVALID_ARGUMENT -7
#define CDS_ERR_READ_FRAME       -8
#define CDS_ERR_CONTROL_NOT_SUPPORTED -9
#define CDS_ERR_BUF_NULL         -10
#define CDS_ERR_BUF_TOO_SMALL    -11
#define CDS_ERR_CONTROL_IO       -12
#define CDS_ERR_UNKNOWN          -512

	// Camera control property IDs match DirectShow's VideoProcAmpProperty /
	// CameraControlProperty values so callers do not need Windows headers.
#define CDS_CONTROL_FLAG_AUTO    0x0001
#define CDS_CONTROL_FLAG_MANUAL  0x0002

#define CDS_VIDEO_PROCAMP_BRIGHTNESS             0
#define CDS_VIDEO_PROCAMP_CONTRAST               1
#define CDS_VIDEO_PROCAMP_HUE                    2
#define CDS_VIDEO_PROCAMP_SATURATION             3
#define CDS_VIDEO_PROCAMP_SHARPNESS              4
#define CDS_VIDEO_PROCAMP_GAMMA                  5
#define CDS_VIDEO_PROCAMP_COLORENABLE            6
#define CDS_VIDEO_PROCAMP_WHITEBALANCE           7
#define CDS_VIDEO_PROCAMP_BACKLIGHTCOMPENSATION  8
#define CDS_VIDEO_PROCAMP_GAIN                   9

#define CDS_CAMERA_CONTROL_PAN                   0
#define CDS_CAMERA_CONTROL_TILT                  1
#define CDS_CAMERA_CONTROL_ROLL                  2
#define CDS_CAMERA_CONTROL_ZOOM                  3
#define CDS_CAMERA_CONTROL_EXPOSURE              4
#define CDS_CAMERA_CONTROL_IRIS                  5
#define CDS_CAMERA_CONTROL_FOCUS                 6

	// Returns CDS_OK when DirectShow initializes successfully, including when
	// zero video devices are currently connected (cds_devices_count returns 0).
	SP_API cds_result_t SP_CALL cds_initialize(void);
	SP_API void         SP_CALL cds_shutdown_capture_api(void);
	SP_API void         SP_CALL cds_set_log_enabled(int32_t enabled); // 0=off, non-zero=on

	// Devices
	SP_API int32_t SP_CALL cds_devices_count(void);

	SP_API size_t  SP_CALL cds_device_name(int32_t device_index, char* buf, size_t buf_len);
	SP_API size_t  SP_CALL cds_device_unique_id(int32_t device_index, char* buf, size_t buf_len); // DevicePath UTF-8
	SP_API size_t  SP_CALL cds_device_model_id(int32_t device_index, char* buf, size_t buf_len);

	SP_API int32_t SP_CALL cds_device_vid(int32_t device_index); // 0 if unknown
	SP_API int32_t SP_CALL cds_device_pid(int32_t device_index); // 0 if unknown

	// Formats (deduped, stable-sorted)
	SP_API int32_t  SP_CALL cds_device_formats_count(int32_t device_index);
	SP_API uint32_t SP_CALL cds_device_format_width(int32_t device_index, int32_t format_index);
	SP_API uint32_t SP_CALL cds_device_format_height(int32_t device_index, int32_t format_index);

	// MAX fps for this format
	SP_API uint32_t SP_CALL cds_device_format_frame_rate(int32_t device_index, int32_t format_index);

	// subtype name: "MJPG","YUY2","NV12","RGB24","RGB32", or GUID string
	SP_API size_t   SP_CALL cds_device_format_type(int32_t device_index, int32_t format_index, char* buf, size_t buf_len);

	// Capture (RGB32 guaranteed, top-down guaranteed)
	SP_API cds_result_t SP_CALL cds_start_capture(uint32_t device_index, uint32_t width, uint32_t height);
	SP_API cds_result_t SP_CALL cds_start_capture_with_format(uint32_t device_index, uint32_t format_index);
	SP_API cds_result_t SP_CALL cds_stop_capture(uint32_t device_index);

	SP_API int32_t      SP_CALL cds_has_first_frame(uint32_t device_index);
	SP_API cds_result_t SP_CALL cds_grab_frame(uint32_t device_index, uint8_t* buffer, size_t available_bytes);

	SP_API int32_t SP_CALL cds_frame_width(uint32_t device_index);
	SP_API int32_t SP_CALL cds_frame_height(uint32_t device_index);
	SP_API int32_t SP_CALL cds_frame_bytes_per_row(uint32_t device_index);

	// VideoProcAmp controls (brightness, contrast, hue, saturation, sharpness,
	// gamma, color enable, white balance, backlight compensation, gain)
	SP_API cds_result_t SP_CALL cds_get_video_proc_amp_range(
		int32_t device_index,
		int32_t property,
		int32_t* min_value,
		int32_t* max_value,
		int32_t* step,
		int32_t* default_value,
		int32_t* caps_flags);
	SP_API cds_result_t SP_CALL cds_get_video_proc_amp(
		int32_t device_index,
		int32_t property,
		int32_t* value,
		int32_t* flags);
	SP_API cds_result_t SP_CALL cds_set_video_proc_amp(
		int32_t device_index,
		int32_t property,
		int32_t value,
		int32_t flags);

	// CameraControl controls (pan, tilt, roll, zoom, exposure, iris, focus)
	SP_API cds_result_t SP_CALL cds_get_camera_control_range(
		int32_t device_index,
		int32_t property,
		int32_t* min_value,
		int32_t* max_value,
		int32_t* step,
		int32_t* default_value,
		int32_t* caps_flags);
	SP_API cds_result_t SP_CALL cds_get_camera_control(
		int32_t device_index,
		int32_t property,
		int32_t* value,
		int32_t* flags);
	SP_API cds_result_t SP_CALL cds_set_camera_control(
		int32_t device_index,
		int32_t property,
		int32_t value,
		int32_t flags);

	// Button press detection WHILE STREAMING (integrated into cds session)
	SP_API int32_t  SP_CALL cds_button_pressed(uint32_t device_index);     // returns 1 once per press (edge), then 0
	SP_API uint64_t SP_CALL cds_button_timestamp(uint32_t device_index);   // timestamp_100ns for last press (best-effort)

#ifdef __cplusplus
}
#endif

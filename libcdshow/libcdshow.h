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
#define CDS_ERR_READ_FRAME       -8
#define CDS_ERR_BUF_NULL         -10
#define CDS_ERR_BUF_TOO_SMALL    -11
#define CDS_ERR_UNKNOWN          -512

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

	// Button press detection WHILE STREAMING (integrated into cds session)
	SP_API int32_t  SP_CALL cds_button_pressed(uint32_t device_index);     // returns 1 once per press (edge), then 0
	SP_API uint64_t SP_CALL cds_button_timestamp(uint32_t device_index);   // timestamp_100ns for last press (best-effort)

#ifdef __cplusplus
}
#endif

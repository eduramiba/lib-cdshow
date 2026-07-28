# Introduction

Very simple library for webcam native video capture that uses DirectShow to export basic C functions so they can be called from any language such as Java, Python, etc.

This was built to be used with JNA in https://github.com/eduramiba/webcam-capture-driver-native

Note: this library has been mostly coded with OpenAI Codex

## Automatic capture-format selection

`cds_start_capture(device, width, height)` selects among native camera modes at
the requested resolution using this policy:

1. Highest advertised frame rate.
2. Lowest expected conversion/decode cost to the library's RGB32 output:
   RGB32, RGB24, NV12, YUY2, MJPG, ARGB32, then unknown subtypes.
3. Lowest stable format index when the scores are otherwise identical.

The selected capability is explicitly configured with its fastest advertised
frame interval. If DirectShow cannot build or run that candidate's RGB32 graph,
resolution-only start tries the next ranked mode at the same resolution. Call
`cds_start_capture_with_format` when the caller needs an exact enumerated mode
without automatic fallback.

## Native NV12/YUY2 output

The existing start functions remain backward-compatible and always request
top-down RGB32/BGRA output. Consumers that can process YUV may use:

- `cds_start_capture_with_output`
- `cds_start_capture_with_format_output`

Pass `CDS_OUTPUT_NATIVE` to preserve an advertised NV12 or YUY2 camera mode
without requesting DirectShow color conversion. Other native subtypes return
`CDS_ERR_FORMAT_NOT_SUPPORTED`; callers can retry the same mode with
`CDS_OUTPUT_BGRA`.

After a successful start, inspect the actual connected frame contract with:

- `cds_frame_pixel_format` / `cds_frame_pixel_format_name`
- `cds_frame_data_size`
- `cds_frame_plane_count`
- `cds_frame_plane_offset`
- `cds_frame_plane_bytes_per_row`

NV12 is exposed as a Y plane followed by an interleaved UV plane. YUY2 is one
packed `Y0 U Y1 V` plane. `cds_grab_frame` copies exactly
`cds_frame_data_size` bytes.

Run `python native_output_smoke.py` with a connected camera to verify direct
YUV capture and the BGRA fallback on the same enumerated mode.

## Camera controls

The C API now exposes DirectShow camera controls through these functions:

- `cds_get_video_proc_amp_range`
- `cds_get_video_proc_amp`
- `cds_set_video_proc_amp`
- `cds_get_camera_control_range`
- `cds_get_camera_control`
- `cds_set_camera_control`

The property IDs are exported in `libcdshow.h` so callers do not need Windows
headers. `VideoProcAmp` covers:

- brightness
- contrast
- hue
- saturation
- sharpness
- gamma
- color enable
- white balance
- backlight compensation
- gain

`CameraControl` covers:

- pan
- tilt
- roll
- zoom
- exposure
- iris
- focus

Use the `*_range` functions first to detect whether a property is supported and
to retrieve its min/max/step/default/capability flags. The capability flags use
`CDS_CONTROL_FLAG_AUTO` and `CDS_CONTROL_FLAG_MANUAL`.

When a device is already streaming, control requests are routed through the live
capture filter. When it is idle, the library opens the device temporarily to
query or set the property.

# Introduction

Very simple library for webcam native video capture that uses DirectShow to export basic C functions so they can be called from any language such as Java, Python, etc.

This was built to be used with JNA in https://github.com/eduramiba/webcam-capture-driver-native

Note: this library has been mostly coded with OpenAI Codex

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

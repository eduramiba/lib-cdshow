import argparse
import ctypes
import os
import sys
import threading
import time


CDS_OK = 0
CDS_OUTPUT_BGRA = 0
CDS_OUTPUT_NATIVE = 1
CDS_PIXEL_FORMAT_BGRA = 1
CDS_PIXEL_FORMAT_NV12 = 2
CDS_PIXEL_FORMAT_YUY2 = 3
CDS_PIXEL_FORMAT_MJPEG = 4
FRAME_CALLBACK = ctypes.WINFUNCTYPE(
    None,
    ctypes.c_uint32,
    ctypes.POINTER(ctypes.c_uint8),
    ctypes.c_size_t,
    ctypes.c_int32,
    ctypes.c_void_p,
)


def configure(dll):
    dll.cds_initialize.restype = ctypes.c_int32
    dll.cds_shutdown_capture_api.restype = None
    dll.cds_devices_count.restype = ctypes.c_int32
    dll.cds_device_formats_count.restype = ctypes.c_int32
    dll.cds_device_format_width.restype = ctypes.c_uint32
    dll.cds_device_format_height.restype = ctypes.c_uint32
    dll.cds_device_format_frame_rate.restype = ctypes.c_uint32
    dll.cds_device_format_type.restype = ctypes.c_size_t
    dll.cds_device_format_type.argtypes = [
        ctypes.c_int32,
        ctypes.c_int32,
        ctypes.c_char_p,
        ctypes.c_size_t,
    ]
    dll.cds_start_capture_with_format_output.restype = ctypes.c_int32
    dll.cds_start_capture_with_format_output.argtypes = [
        ctypes.c_uint32,
        ctypes.c_uint32,
        ctypes.c_int32,
    ]
    dll.cds_stop_capture.restype = ctypes.c_int32
    dll.cds_has_first_frame.restype = ctypes.c_int32
    dll.cds_frame_width.restype = ctypes.c_int32
    dll.cds_frame_height.restype = ctypes.c_int32
    dll.cds_frame_bytes_per_row.restype = ctypes.c_int32
    dll.cds_frame_data_size.restype = ctypes.c_int32
    dll.cds_frame_pixel_format.restype = ctypes.c_int32
    dll.cds_frame_plane_count.restype = ctypes.c_int32
    dll.cds_frame_plane_offset.restype = ctypes.c_int32
    dll.cds_frame_plane_bytes_per_row.restype = ctypes.c_int32
    dll.cds_grab_frame.restype = ctypes.c_int32
    dll.cds_grab_frame.argtypes = [
        ctypes.c_uint32,
        ctypes.POINTER(ctypes.c_uint8),
        ctypes.c_size_t,
    ]
    dll.cds_set_frame_callback.restype = ctypes.c_int32
    dll.cds_set_frame_callback.argtypes = [
        ctypes.c_uint32,
        FRAME_CALLBACK,
        ctypes.c_void_p,
        ctypes.c_int32,
    ]


def format_type(dll, device_index, format_index):
    buffer = ctypes.create_string_buffer(128)
    dll.cds_device_format_type(device_index, format_index, buffer, len(buffer))
    return buffer.value.decode("ascii", errors="replace").upper()


def wait_for_frame(dll, device_index, timeout_seconds):
    deadline = time.monotonic() + timeout_seconds
    while time.monotonic() < deadline:
        if dll.cds_has_first_frame(device_index) > 0:
            return
        time.sleep(0.02)
    raise RuntimeError("Timed out waiting for the first frame")


def assert_layout_and_grab(dll, device_index, expected_pixel_format):
    width = dll.cds_frame_width(device_index)
    height = dll.cds_frame_height(device_index)
    data_size = dll.cds_frame_data_size(device_index)
    pixel_format = dll.cds_frame_pixel_format(device_index)
    plane_count = dll.cds_frame_plane_count(device_index)

    assert width > 0 and height > 0
    assert data_size > 0
    assert pixel_format == expected_pixel_format
    assert plane_count == (2 if expected_pixel_format == CDS_PIXEL_FORMAT_NV12 else 1)

    offsets = [
        dll.cds_frame_plane_offset(device_index, plane)
        for plane in range(plane_count)
    ]
    strides = [
        dll.cds_frame_plane_bytes_per_row(device_index, plane)
        for plane in range(plane_count)
    ]
    assert offsets[0] == 0
    if expected_pixel_format == CDS_PIXEL_FORMAT_MJPEG:
        assert strides == [0]
    else:
        assert all(stride > 0 for stride in strides)

    if expected_pixel_format == CDS_PIXEL_FORMAT_BGRA:
        assert strides[0] >= width * 4
        assert data_size >= strides[0] * height
    elif expected_pixel_format == CDS_PIXEL_FORMAT_YUY2:
        assert width % 2 == 0
        assert strides[0] >= width * 2
        assert data_size >= strides[0] * height
    elif expected_pixel_format == CDS_PIXEL_FORMAT_NV12:
        assert width % 2 == 0 and height % 2 == 0
        assert strides[0] >= width and strides[1] >= width
        assert offsets[1] >= strides[0] * height
        assert data_size >= offsets[1] + strides[1] * (height // 2)

    frame = None
    if expected_pixel_format != CDS_PIXEL_FORMAT_MJPEG:
        frame = (ctypes.c_uint8 * data_size)()
        result = dll.cds_grab_frame(device_index, frame, data_size)
        assert result == CDS_OK, f"cds_grab_frame returned {result}"
        assert any(frame), "Captured frame was entirely zero"
    return {
        "width": width,
        "height": height,
        "data_size": data_size,
        "plane_offsets": offsets,
        "plane_strides": strides,
    }


def main():
    parser = argparse.ArgumentParser()
    parser.add_argument(
        "--dll",
        default=os.path.join("x64", "Release", "libcdshow.dll"),
    )
    parser.add_argument("--timeout", type=float, default=10.0)
    args = parser.parse_args()

    dll_path = os.path.abspath(args.dll)
    dll = ctypes.WinDLL(dll_path)
    configure(dll)

    initialized = dll.cds_initialize()
    if initialized != CDS_OK:
        raise RuntimeError(f"cds_initialize returned {initialized}")

    try:
        candidates = []
        for device_index in range(dll.cds_devices_count()):
            for format_index in range(dll.cds_device_formats_count(device_index)):
                subtype = format_type(dll, device_index, format_index)
                if subtype not in {"NV12", "YUY2", "MJPG"}:
                    continue
                candidates.append(
                    (
                        dll.cds_device_format_frame_rate(device_index, format_index),
                        dll.cds_device_format_width(device_index, format_index)
                        * dll.cds_device_format_height(device_index, format_index),
                        device_index,
                        format_index,
                        subtype,
                    )
                )

        if not candidates:
            print("SKIP: No connected camera advertises NV12, YUY2, or MJPG")
            return 0

        _, _, device_index, format_index, subtype = max(candidates)
        expected_native = {
            "NV12": CDS_PIXEL_FORMAT_NV12,
            "YUY2": CDS_PIXEL_FORMAT_YUY2,
            "MJPG": CDS_PIXEL_FORMAT_MJPEG,
        }[subtype]

        result = dll.cds_start_capture_with_format_output(
            device_index,
            format_index,
            CDS_OUTPUT_NATIVE,
        )
        assert result == CDS_OK, f"Native start returned {result}"
        try:
            wait_for_frame(dll, device_index, args.timeout)
            callback_event = threading.Event()
            callback_frames = 0
            callback_bytes = 0
            callback_jpeg = False

            @FRAME_CALLBACK
            def on_frame(callback_device, data, data_size, bottom_up, _user_data):
                nonlocal callback_frames, callback_bytes, callback_jpeg
                assert callback_device == device_index
                assert bool(data)
                assert data_size > 0
                assert bottom_up == 0
                callback_frames += 1
                callback_bytes += data_size
                if expected_native == CDS_PIXEL_FORMAT_MJPEG:
                    callback_jpeg = bytes(data[:2]) == b"\xff\xd8"
                if callback_frames >= 5:
                    callback_event.set()

            callback_result = dll.cds_set_frame_callback(
                device_index,
                on_frame,
                None,
                1,
            )
            assert callback_result == CDS_OK
            assert callback_event.wait(args.timeout), "Timed out waiting for callback frames"
            native_layout = assert_layout_and_grab(
                dll,
                device_index,
                expected_native,
            )
            assert callback_frames >= 5
            assert callback_bytes > 0
            if expected_native == CDS_PIXEL_FORMAT_MJPEG:
                assert callback_jpeg, "MJPEG callback sample is not a JPEG"
            assert dll.cds_set_frame_callback(
                device_index,
                FRAME_CALLBACK(),
                None,
                0,
            ) == CDS_OK
            native_layout["callback_frames"] = callback_frames
        finally:
            dll.cds_stop_capture(device_index)

        result = dll.cds_start_capture_with_format_output(
            device_index,
            format_index,
            CDS_OUTPUT_BGRA,
        )
        assert result == CDS_OK, f"BGRA fallback start returned {result}"
        try:
            wait_for_frame(dll, device_index, args.timeout)
            bgra_layout = assert_layout_and_grab(
                dll,
                device_index,
                CDS_PIXEL_FORMAT_BGRA,
            )
        finally:
            dll.cds_stop_capture(device_index)

        print(
            "PASS:",
            {
                "device_index": device_index,
                "format_index": format_index,
                "native_subtype": subtype,
                "native": native_layout,
                "bgra": bgra_layout,
            },
        )
        return 0
    finally:
        dll.cds_shutdown_capture_api()


if __name__ == "__main__":
    try:
        sys.exit(main())
    except Exception as error:
        print(f"FAIL: {error}", file=sys.stderr)
        raise

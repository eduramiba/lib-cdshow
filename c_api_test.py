import ctypes
import os
import time
import platform
from ctypes import wintypes
from PIL import Image

print("Python bitness:", platform.architecture())
print("Pointer size:", ctypes.sizeof(ctypes.c_void_p))

DLL_PATH = os.path.abspath("x64/Release/libcdshow.dll")
dll = ctypes.WinDLL(DLL_PATH)  # stdcall

CDS_OK = 0

# ==========================
# Prototypes
# ==========================

dll.cds_initialize.restype = ctypes.c_int32
dll.cds_shutdown_capture_api.restype = None
dll.cds_set_log_enabled.restype = None
dll.cds_set_log_enabled.argtypes = [ctypes.c_int32]

dll.cds_devices_count.restype = ctypes.c_int32

dll.cds_device_name.restype = ctypes.c_size_t
dll.cds_device_name.argtypes = [ctypes.c_int32, ctypes.c_char_p, ctypes.c_size_t]
dll.cds_device_unique_id.restype = ctypes.c_size_t
dll.cds_device_unique_id.argtypes = [ctypes.c_int32, ctypes.c_char_p, ctypes.c_size_t]

dll.cds_device_formats_count.restype = ctypes.c_int32
dll.cds_device_format_width.restype = ctypes.c_uint32
dll.cds_device_format_height.restype = ctypes.c_uint32
dll.cds_device_format_frame_rate.restype = ctypes.c_uint32

dll.cds_start_capture_with_format.restype = ctypes.c_int32
dll.cds_stop_capture.restype = ctypes.c_int32

dll.cds_has_first_frame.restype = ctypes.c_int32
dll.cds_frame_width.restype = ctypes.c_int32
dll.cds_frame_height.restype = ctypes.c_int32

dll.cds_grab_frame.restype = ctypes.c_int32
dll.cds_grab_frame.argtypes = [ctypes.c_uint32, ctypes.POINTER(ctypes.c_uint8), ctypes.c_size_t]

dll.cds_button_pressed.restype = ctypes.c_int32
dll.cds_button_timestamp.restype = ctypes.c_uint64

# ==========================
# Helpers
# ==========================

def get_device_name(index):
    buf = ctypes.create_string_buffer(512)
    dll.cds_device_name(index, buf, 512)
    return buf.value.decode("utf-8")

def get_device_unique_id(index):
    buf = ctypes.create_string_buffer(1024)
    dll.cds_device_unique_id(index, buf, 1024)
    return buf.value.decode("utf-8")

def pick_device_index():
    count = dll.cds_devices_count()
    if count <= 0:
        return -1

    devices = [(i, get_device_name(i)) for i in range(count)]

    non_fhd = [
        i for i, name in devices
        if name.strip().lower() != "fhd webcam"
    ]

    if non_fhd:
        return non_fhd[0]
    return devices[0][0]

# ==========================
# Main
# ==========================

def main():
    dll.cds_set_log_enabled(1)

    if dll.cds_initialize() != CDS_OK:
        print("cds_initialize failed")
        return

    count = dll.cds_devices_count()
    if count <= 0:
        print("No camera found")
        dll.cds_shutdown_capture_api()
        return

    dev_index = pick_device_index()
    if dev_index < 0:
        print("No camera found")
        dll.cds_shutdown_capture_api()
        return

    print("Using device:", get_device_name(dev_index))
    print("Device ID:", get_device_unique_id(dev_index))

    fmt_count = dll.cds_device_formats_count(dev_index)

    best_fmt = 0
    best_pixels = 0
    best_fps = 0

    for i in range(fmt_count):
        w = dll.cds_device_format_width(dev_index, i)
        h = dll.cds_device_format_height(dev_index, i)
        fps = dll.cds_device_format_frame_rate(dev_index, i)

        pixels = w * h
        if pixels > best_pixels or (pixels == best_pixels and fps > best_fps):
            best_fmt = i
            best_pixels = pixels
            best_fps = fps

    if dll.cds_start_capture_with_format(dev_index, best_fmt) != CDS_OK:
        print("Failed to start capture")
        dll.cds_shutdown_capture_api()
        return

    print("Waiting for first frame...")
    while not dll.cds_has_first_frame(dev_index):
        time.sleep(0.05)

    width = dll.cds_frame_width(dev_index)
    height = dll.cds_frame_height(dev_index)

    print("Streaming:", width, "x", height)

    frame_size = width * height * 4
    buffer = (ctypes.c_uint8 * frame_size)()

    frame_counter = 0

    try:
        while True:
            if dll.cds_button_pressed(dev_index):
                ts = dll.cds_button_timestamp(dev_index)
                print(f"[BUTTON PRESSED] ts100ns={ts}")

                rc = dll.cds_grab_frame(dev_index, buffer, frame_size)
                if rc != CDS_OK:
                    print("Frame error:", rc)
                    time.sleep(0.1)
                    continue

                img = Image.frombuffer(
                    "RGB",
                    (width, height),
                    buffer,
                    "raw",
                    "BGRX",
                    0,
                    1
                )

                filename = f"frame_{frame_counter:05d}.jpg"
                img.save(filename, "JPEG", quality=90)
                print("Saved", filename)

                frame_counter += 1

            time.sleep(0.01)

    except KeyboardInterrupt:
        print("Stopping...")

    finally:
        dll.cds_stop_capture(dev_index)
        dll.cds_shutdown_capture_api()


if __name__ == "__main__":
    main()

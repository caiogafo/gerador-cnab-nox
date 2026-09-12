"""Capture an owned X11/Tk window directly, without reading the desktop root."""

from __future__ import annotations

import ctypes as c
import ctypes.util
import sys

from PIL import Image


class XImage(c.Structure):
    _fields_ = [
        ("width", c.c_int),
        ("height", c.c_int),
        ("xoffset", c.c_int),
        ("format", c.c_int),
        ("data", c.c_void_p),
        ("byte_order", c.c_int),
        ("bitmap_unit", c.c_int),
        ("bitmap_bit_order", c.c_int),
        ("bitmap_pad", c.c_int),
        ("depth", c.c_int),
        ("bytes_per_line", c.c_int),
        ("bits_per_pixel", c.c_int),
        ("red_mask", c.c_ulong),
        ("green_mask", c.c_ulong),
        ("blue_mask", c.c_ulong),
    ]


def capture(window, width, height, destination):
    lib = c.CDLL(ctypes.util.find_library("X11"))
    lib.XOpenDisplay.argtypes = [c.c_char_p]
    lib.XOpenDisplay.restype = c.c_void_p
    lib.XGetImage.argtypes = [
        c.c_void_p,
        c.c_ulong,
        c.c_int,
        c.c_int,
        c.c_uint,
        c.c_uint,
        c.c_ulong,
        c.c_int,
    ]
    lib.XGetImage.restype = c.POINTER(XImage)
    lib.XDestroyImage.argtypes = [c.POINTER(XImage)]
    lib.XCloseDisplay.argtypes = [c.c_void_p]
    display = lib.XOpenDisplay(None)
    if not display:
        raise RuntimeError("X11 display unavailable")
    pointer = lib.XGetImage(display, window, 0, 0, width, height, c.c_ulong(-1).value, 2)
    if not pointer:
        raise RuntimeError("Owned window image unavailable")
    try:
        img = pointer.contents
        if (img.bits_per_pixel, img.byte_order, img.red_mask, img.green_mask, img.blue_mask) != (
            32,
            0,
            0xFF0000,
            0xFF00,
            0xFF,
        ):
            raise RuntimeError("Unsupported pixel format")
        data = c.string_at(img.data, img.bytes_per_line * height)
        Image.frombytes("RGB", (width, height), data, "raw", "BGRX", img.bytes_per_line).save(
            destination
        )
    finally:
        lib.XDestroyImage(pointer)
        lib.XCloseDisplay(display)


if __name__ == "__main__":
    capture(int(sys.argv[1]), int(sys.argv[2]), int(sys.argv[3]), sys.argv[4])

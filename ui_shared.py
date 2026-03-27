from __future__ import annotations

import ctypes
import os
from tkinter import TclError


STYLE = {
    "primario": "#ECD925",
    "secundario": "#282828",
    "exito": "#008D53",
    "advertencia": "#ff1500",
    "peligro": "#d74a3d",
    "fondo": "#F8F9FA",
    "surface": "#F8F9FA",
    "texto_oscuro": "#282828",
    "texto_claro": "#ffffff",
    "borde": "#F8F9FA",
}

FONT_TITLE = ("Inter", 22, "bold")
FONT_SUBTITLE = ("Inter", 17, "bold")
FONT_LABEL = ("Inter", 15)
FONT_SMALL = ("Inter", 14)

BASE_FONTS = {
    "title": FONT_TITLE,
    "subtitle": FONT_SUBTITLE,
    "label": FONT_LABEL,
    "label_bold": ("Inter", 15, "bold"),
    "small": FONT_SMALL,
    "small_bold": ("Inter", 14, "bold"),
}

FONTS = dict(BASE_FONTS)


_MONITOR_DEFAULTTONEAREST = 2


class _WinRect(ctypes.Structure):
    _fields_ = [
        ("left", ctypes.c_long),
        ("top", ctypes.c_long),
        ("right", ctypes.c_long),
        ("bottom", ctypes.c_long),
    ]


class _MonitorInfo(ctypes.Structure):
    _fields_ = [
        ("cbSize", ctypes.c_ulong),
        ("rcMonitor", _WinRect),
        ("rcWork", _WinRect),
        ("dwFlags", ctypes.c_ulong),
    ]


def _scaled_font(font_value: tuple, factor: float) -> tuple:
    family = font_value[0]
    size = int(font_value[1])
    traits = tuple(font_value[2:])
    scaled_size = max(9, int(round(size * factor)))
    return (family, scaled_size, *traits)


def _get_window_work_area(window) -> tuple[int, int, int, int] | None:
    if os.name != "nt" or window is None:
        return None

    try:
        hwnd = int(window.winfo_id())
    except (AttributeError, TclError, ValueError):
        return None

    try:
        user32 = ctypes.windll.user32
        monitor = user32.MonitorFromWindow(hwnd, _MONITOR_DEFAULTTONEAREST)
        if not monitor:
            return None

        monitor_info = _MonitorInfo()
        monitor_info.cbSize = ctypes.sizeof(_MonitorInfo)
        if not user32.GetMonitorInfoW(monitor, ctypes.byref(monitor_info)):
            return None

        work_area = monitor_info.rcWork
        return work_area.left, work_area.top, work_area.right, work_area.bottom
    except (AttributeError, OSError):
        return None


def _position_toplevel(window, parent, width: int, height: int) -> None:
    try:
        window.update_idletasks()
    except TclError:
        return

    work_area = _get_window_work_area(parent or window)
    if work_area is None:
        left = 0
        top = 0
        right = max(width, int(window.winfo_screenwidth()))
        bottom = max(height, int(window.winfo_screenheight()))
    else:
        left, top, right, bottom = work_area

    x = left + max(0, ((right - left) - width) // 2)
    y = top + max(0, ((bottom - top) - height) // 2)

    if parent is not None:
        try:
            if parent.winfo_exists():
                parent.update_idletasks()
                parent_width = max(parent.winfo_width(), parent.winfo_reqwidth())
                parent_height = max(parent.winfo_height(), parent.winfo_reqheight())
                if parent_width > 1 and parent_height > 1:
                    x = parent.winfo_rootx() + (parent_width - width) // 2
                    y = parent.winfo_rooty() + (parent_height - height) // 2
        except TclError:
            pass

    padding = 16
    max_x = max(left + padding, right - width - padding)
    max_y = max(top + padding, bottom - height - padding)
    x = min(max(left + padding, x), max_x)
    y = min(max(top + padding, y), max_y)
    window.geometry(f"{width}x{height}+{x}+{y}")


def _safe_focus(widget) -> None:
    def _apply_focus() -> None:
        try:
            if widget is not None and widget.winfo_exists() and widget.winfo_toplevel().winfo_exists():
                widget.focus_set()
        except TclError:
            return

    try:
        if widget is not None and widget.winfo_exists():
            widget.after(20, _apply_focus)
    except TclError:
        return

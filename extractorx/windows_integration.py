"""Native Windows shell hooks for Tkinter.

This module keeps platform-specific ctypes code out of the UI layer. It provides
two legacy-parity features without third-party dependencies:

- WM_DROPFILES drag/drop support
- Shell notification-area icon support
"""

from __future__ import annotations

import os
import ctypes
from ctypes import wintypes
from typing import Callable


IS_WINDOWS = os.name == "nt"

if IS_WINDOWS:
    user32 = ctypes.windll.user32
    shell32 = ctypes.windll.shell32
else:  # pragma: no cover - only exercised on non-Windows development hosts.
    user32 = None
    shell32 = None


WM_DROPFILES = 0x0233
WM_APP = 0x8000
WM_TRAYICON = WM_APP + 91
WM_LBUTTONDBLCLK = 0x0203
WM_LBUTTONUP = 0x0202
WM_RBUTTONUP = 0x0205
GWL_WNDPROC = -4

NIM_ADD = 0x00000000
NIM_MODIFY = 0x00000001
NIM_DELETE = 0x00000002
NIM_SETVERSION = 0x00000004
NIF_MESSAGE = 0x00000001
NIF_ICON = 0x00000002
NIF_TIP = 0x00000004
NOTIFYICON_VERSION_4 = 4
IDI_APPLICATION = 32512


class NOTIFYICONDATAW(ctypes.Structure):
    _fields_ = [
        ("cbSize", wintypes.DWORD),
        ("hWnd", wintypes.HWND),
        ("uID", wintypes.UINT),
        ("uFlags", wintypes.UINT),
        ("uCallbackMessage", wintypes.UINT),
        ("hIcon", wintypes.HICON),
        ("szTip", wintypes.WCHAR * 128),
        ("dwState", wintypes.DWORD),
        ("dwStateMask", wintypes.DWORD),
        ("szInfo", wintypes.WCHAR * 256),
        ("uVersion", wintypes.UINT),
        ("szInfoTitle", wintypes.WCHAR * 64),
        ("dwInfoFlags", wintypes.DWORD),
        ("guidItem", ctypes.c_byte * 16),
        ("hBalloonIcon", wintypes.HICON),
    ]


if IS_WINDOWS:
    LRESULT = ctypes.c_ssize_t
    WNDPROC = ctypes.WINFUNCTYPE(LRESULT, wintypes.HWND, wintypes.UINT, wintypes.WPARAM, wintypes.LPARAM)

    user32.CallWindowProcW.argtypes = [ctypes.c_void_p, wintypes.HWND, wintypes.UINT, wintypes.WPARAM, wintypes.LPARAM]
    user32.CallWindowProcW.restype = LRESULT
    user32.SetWindowLongPtrW.argtypes = [wintypes.HWND, ctypes.c_int, ctypes.c_void_p]
    user32.SetWindowLongPtrW.restype = ctypes.c_void_p
    user32.LoadIconW.argtypes = [wintypes.HINSTANCE, wintypes.LPCWSTR]
    user32.LoadIconW.restype = wintypes.HICON
    shell32.DragAcceptFiles.argtypes = [wintypes.HWND, wintypes.BOOL]
    shell32.DragQueryFileW.argtypes = [wintypes.HANDLE, wintypes.UINT, wintypes.LPWSTR, wintypes.UINT]
    shell32.DragQueryFileW.restype = wintypes.UINT
    shell32.DragFinish.argtypes = [wintypes.HANDLE]
    shell32.Shell_NotifyIconW.argtypes = [wintypes.DWORD, ctypes.POINTER(NOTIFYICONDATAW)]
    shell32.Shell_NotifyIconW.restype = wintypes.BOOL


def detect_system_theme() -> str:
    """Return ``'dark'`` or ``'light'`` based on the Windows personalization setting."""
    if not IS_WINDOWS:
        return "dark"
    try:
        import winreg
        with winreg.OpenKey(
            winreg.HKEY_CURRENT_USER,
            r"SOFTWARE\Microsoft\Windows\CurrentVersion\Themes\Personalize",
            0,
            winreg.KEY_READ,
        ) as key:
            value, _ = winreg.QueryValueEx(key, "AppsUseLightTheme")
            return "light" if value else "dark"
    except OSError:
        return "dark"


def set_dark_titlebar(hwnd: int, dark: bool = True) -> None:
    """Enable or disable the immersive dark title bar on Windows 10 20H1+ / Windows 11."""
    if not IS_WINDOWS:
        return
    try:
        dwmapi = ctypes.windll.dwmapi
        value = ctypes.c_int(1 if dark else 0)
        # DWMWA_USE_IMMERSIVE_DARK_MODE = 20
        dwmapi.DwmSetWindowAttribute(hwnd, 20, ctypes.byref(value), ctypes.sizeof(value))
    except (OSError, AttributeError):
        pass


# ITaskbarList3 COM interface for taskbar progress
TBPF_NOPROGRESS = 0x0
TBPF_NORMAL = 0x2
TBPF_ERROR = 0x4
TBPF_PAUSED = 0x8

_CLSID_TaskbarList = b"\x44\xf3\xfd\x56\x6d\xfd\xd0\x11\x95\x8a\x00\x60\x97\xc9\xa0\x90"
_IID_ITaskbarList3 = b"\x91\xfb\x1a\xea\x28\x9e\x86\x4b\x90\xe9\x9e\x9f\x8a\x5e\xef\xaf"


class TaskbarProgress:
    """Thin wrapper around ITaskbarList3 for taskbar progress indication."""

    def __init__(self, hwnd: int) -> None:
        self.hwnd = hwnd
        self._taskbar = None
        if IS_WINDOWS:
            self._init_com()

    def _init_com(self) -> None:
        try:
            ole32 = ctypes.windll.ole32
            ole32.CoInitialize(None)
            clsid = (ctypes.c_byte * 16)(*_CLSID_TaskbarList)
            iid = (ctypes.c_byte * 16)(*_IID_ITaskbarList3)
            p = ctypes.c_void_p()
            hr = ole32.CoCreateInstance(
                ctypes.byref(clsid), None, 1, ctypes.byref(iid), ctypes.byref(p)
            )
            if hr == 0 and p.value:
                self._taskbar = p.value
                # Call HrInit (vtable index 3)
                vtable = ctypes.cast(
                    ctypes.cast(p, ctypes.POINTER(ctypes.c_void_p))[0],
                    ctypes.POINTER(ctypes.c_void_p),
                )
                hr_init = ctypes.WINFUNCTYPE(ctypes.HRESULT, ctypes.c_void_p)(vtable[3])
                hr_init(p)
        except (OSError, AttributeError):
            self._taskbar = None

    def set_progress(self, current: int, total: int) -> None:
        if not self._taskbar:
            return
        try:
            vtable = ctypes.cast(
                ctypes.cast(self._taskbar, ctypes.POINTER(ctypes.c_void_p))[0],
                ctypes.POINTER(ctypes.c_void_p),
            )
            # SetProgressValue (vtable index 9)
            func = ctypes.WINFUNCTYPE(
                ctypes.HRESULT, ctypes.c_void_p, ctypes.c_void_p,
                ctypes.c_ulonglong, ctypes.c_ulonglong,
            )(vtable[9])
            func(self._taskbar, self.hwnd, current, total)
        except (OSError, ValueError):
            pass

    def set_state(self, state: int) -> None:
        if not self._taskbar:
            return
        try:
            vtable = ctypes.cast(
                ctypes.cast(self._taskbar, ctypes.POINTER(ctypes.c_void_p))[0],
                ctypes.POINTER(ctypes.c_void_p),
            )
            # SetProgressState (vtable index 10)
            func = ctypes.WINFUNCTYPE(
                ctypes.HRESULT, ctypes.c_void_p, ctypes.c_void_p, ctypes.c_int,
            )(vtable[10])
            func(self._taskbar, self.hwnd, state)
        except (OSError, ValueError):
            pass

    def clear(self) -> None:
        self.set_state(TBPF_NOPROGRESS)


class WindowsShellBridge:
    def __init__(
        self,
        root,
        on_files_dropped: Callable[[list[str]], None],
        on_tray_restore: Callable[[], None],
        on_tray_menu: Callable[[], None],
    ) -> None:
        self.root = root
        self.on_files_dropped = on_files_dropped
        self.on_tray_restore = on_tray_restore
        self.on_tray_menu = on_tray_menu
        self.enabled = IS_WINDOWS
        self.hwnd = int(root.winfo_id()) if self.enabled else 0
        self.old_wndproc: ctypes.c_void_p | None = None
        self.wndproc = None
        self.tray_visible = False
        self.icon_data: NOTIFYICONDATAW | None = None
        if self.enabled:
            self._subclass_window()

    def enable_drag_drop(self) -> None:
        if self.enabled:
            shell32.DragAcceptFiles(self.hwnd, True)

    def add_tray_icon(self, tip: str) -> None:
        if not self.enabled or self.tray_visible:
            return
        icon = user32.LoadIconW(None, ctypes.c_wchar_p(IDI_APPLICATION))
        data = NOTIFYICONDATAW()
        data.cbSize = ctypes.sizeof(NOTIFYICONDATAW)
        data.hWnd = self.hwnd
        data.uID = 1
        data.uFlags = NIF_MESSAGE | NIF_ICON | NIF_TIP
        data.uCallbackMessage = WM_TRAYICON
        data.hIcon = icon
        data.szTip = tip[:127]
        data.uVersion = NOTIFYICON_VERSION_4
        if shell32.Shell_NotifyIconW(NIM_ADD, ctypes.byref(data)):
            shell32.Shell_NotifyIconW(NIM_SETVERSION, ctypes.byref(data))
            self.icon_data = data
            self.tray_visible = True

    def update_tray_tip(self, tip: str) -> None:
        if not self.enabled or not self.tray_visible or not self.icon_data:
            return
        self.icon_data.szTip = tip[:127]
        shell32.Shell_NotifyIconW(NIM_MODIFY, ctypes.byref(self.icon_data))

    def remove_tray_icon(self) -> None:
        if self.enabled and self.tray_visible and self.icon_data:
            shell32.Shell_NotifyIconW(NIM_DELETE, ctypes.byref(self.icon_data))
        self.tray_visible = False

    def dispose(self) -> None:
        if not self.enabled:
            return
        self.remove_tray_icon()
        shell32.DragAcceptFiles(self.hwnd, False)
        if self.old_wndproc:
            user32.SetWindowLongPtrW(self.hwnd, GWL_WNDPROC, self.old_wndproc)
            self.old_wndproc = None

    def _subclass_window(self) -> None:
        if self.old_wndproc:
            return

        @WNDPROC
        def wndproc(hwnd, msg, wparam, lparam):
            if msg == WM_DROPFILES:
                paths = self._read_drop_paths(wparam)
                self.root.after(0, lambda: self.on_files_dropped(paths))
                return 0
            if msg == WM_TRAYICON:
                if int(lparam) in (WM_LBUTTONUP, WM_LBUTTONDBLCLK):
                    self.root.after(0, self.on_tray_restore)
                    return 0
                if int(lparam) == WM_RBUTTONUP:
                    self.root.after(0, self.on_tray_menu)
                    return 0
            return user32.CallWindowProcW(self.old_wndproc, hwnd, msg, wparam, lparam)

        self.wndproc = wndproc
        self.old_wndproc = user32.SetWindowLongPtrW(self.hwnd, GWL_WNDPROC, ctypes.cast(wndproc, ctypes.c_void_p))

    @staticmethod
    def _read_drop_paths(drop_handle: int) -> list[str]:
        count = shell32.DragQueryFileW(drop_handle, 0xFFFFFFFF, None, 0)
        paths: list[str] = []
        try:
            for index in range(count):
                length = shell32.DragQueryFileW(drop_handle, index, None, 0)
                buffer = ctypes.create_unicode_buffer(length + 1)
                shell32.DragQueryFileW(drop_handle, index, buffer, length + 1)
                paths.append(buffer.value)
        finally:
            shell32.DragFinish(drop_handle)
        return paths

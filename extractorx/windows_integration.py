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

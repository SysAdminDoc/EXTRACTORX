"""Windows Explorer context menu integration for the Python port."""

from __future__ import annotations

import os
import shutil
import sys
from pathlib import Path

from .archive import ARCHIVE_EXTENSIONS

try:
    import winreg
except ImportError:  # pragma: no cover - this module is Windows-only at runtime.
    winreg = None  # type: ignore[assignment]


MENU_NAME = "ExtractorX"


def is_supported() -> bool:
    return os.name == "nt" and winreg is not None


def install_context_menu(config: dict, entrypoint: Path | None = None) -> None:
    if not is_supported():
        raise OSError("Explorer integration is only available on Windows.")
    uninstall_context_menu()
    entrypoint = entrypoint or Path(__file__).resolve().parents[1] / "ExtractorX.py"
    grouped = bool(config.get("CtxGrouped", True))

    for extension in sorted(ARCHIVE_EXTENSIONS):
        base = rf"Software\Classes\SystemFileAssociations\{extension}\shell"
        if grouped:
            group_key = rf"{base}\{MENU_NAME}"
            _set_value(group_key, "", MENU_NAME)
            _set_value(group_key, "MUIVerb", MENU_NAME)
            _set_value(group_key, "SubCommands", "")
            if bool(config.get("CtxExtractHere", True)):
                _set_command(rf"{group_key}\shell\01_here", "Extract here", _command(entrypoint, "--extract-here", "--auto-extract"))
            if bool(config.get("CtxExtractToFolder", True)):
                _set_command(rf"{group_key}\shell\02_folder", "Extract to folder", _command(entrypoint, "--auto-extract"))
            if bool(config.get("CtxEnqueue", True)):
                _set_command(rf"{group_key}\shell\03_enqueue", "Add to ExtractorX", _command(entrypoint))
        else:
            if bool(config.get("CtxExtractHere", True)):
                _set_command(rf"{base}\{MENU_NAME}_Here", "ExtractorX: Extract here", _command(entrypoint, "--extract-here", "--auto-extract"))
            if bool(config.get("CtxExtractToFolder", True)):
                _set_command(rf"{base}\{MENU_NAME}_Folder", "ExtractorX: Extract to folder", _command(entrypoint, "--auto-extract"))
            if bool(config.get("CtxEnqueue", True)):
                _set_command(rf"{base}\{MENU_NAME}_Enqueue", "ExtractorX: Add to queue", _command(entrypoint))

    if bool(config.get("CtxSearchArchives", True)):
        _set_command(rf"Software\Classes\Directory\shell\{MENU_NAME}", "Search for archives with ExtractorX", _command(entrypoint, "--scan"))
        _set_command(rf"Software\Classes\Directory\Background\shell\{MENU_NAME}", "Search for archives with ExtractorX", _command(entrypoint, "--scan", "%V"))


def install_file_associations(extensions: list[str], entrypoint: Path | None = None) -> None:
    if not is_supported():
        raise OSError("File associations are only available on Windows.")
    entrypoint = entrypoint or Path(__file__).resolve().parents[1] / "ExtractorX.py"
    for extension in _normalize_extensions(extensions):
        prog_id = _association_prog_id(extension)
        # Back up any existing default ProgID before we claim the extension,
        # so uninstall can restore the user's original handler (e.g. 7-Zip).
        previous = _read_default(rf"Software\Classes\{extension}")
        if previous and previous != prog_id:
            _set_value(rf"Software\Classes\{prog_id}", "BackupProgID", previous)
        _set_value(rf"Software\Classes\{extension}", "", prog_id)
        _set_value(rf"Software\Classes\{prog_id}", "", f"ExtractorX archive ({extension})")
        _set_command(rf"Software\Classes\{prog_id}\shell\open", "Open with ExtractorX", _command(entrypoint))


def uninstall_file_associations(extensions: list[str]) -> None:
    if not is_supported():
        return
    for extension in _normalize_extensions(extensions):
        prog_id = _association_prog_id(extension)
        backup = _read_value(rf"Software\Classes\{prog_id}", "BackupProgID")
        try:
            with winreg.OpenKey(winreg.HKEY_CURRENT_USER, rf"Software\Classes\{extension}", 0, winreg.KEY_READ | winreg.KEY_WRITE) as key:
                current, _ = winreg.QueryValueEx(key, "")
                if current == prog_id:
                    if backup:
                        winreg.SetValueEx(key, "", 0, winreg.REG_SZ, backup)
                    else:
                        winreg.DeleteValue(key, "")
        except OSError:
            pass
        _delete_tree(rf"Software\Classes\{prog_id}")


def uninstall_context_menu() -> None:
    if not is_supported():
        return
    for extension in sorted(ARCHIVE_EXTENSIONS):
        base = rf"Software\Classes\SystemFileAssociations\{extension}\shell"
        for name in (MENU_NAME, f"{MENU_NAME}_Here", f"{MENU_NAME}_Folder", f"{MENU_NAME}_Enqueue"):
            _delete_tree(rf"{base}\{name}")
    _delete_tree(rf"Software\Classes\Directory\shell\{MENU_NAME}")
    _delete_tree(rf"Software\Classes\Directory\Background\shell\{MENU_NAME}")


def _command(entrypoint: Path, *flags: str) -> str:
    python = shutil.which("pythonw.exe") or shutil.which("python.exe") or sys.executable
    parts = [f'"{python}"', f'"{entrypoint}"']
    parts.extend(flag for flag in flags if flag != "%V")
    if "%V" in flags:
        parts.append('"%V"')
    else:
        parts.append('"%1"')
    return " ".join(parts)


def _association_prog_id(extension: str) -> str:
    return f"{MENU_NAME}.{extension.strip().lstrip('.')}"


def _normalize_extensions(extensions: list[str]) -> list[str]:
    normalized: list[str] = []
    for extension in extensions:
        value = str(extension).strip().lower()
        if not value:
            continue
        if not value.startswith("."):
            value = "." + value
        if value not in normalized:
            normalized.append(value)
    return normalized


def _set_command(path: str, label: str, command: str) -> None:
    _set_value(path, "", label)
    _set_value(path + r"\command", "", command)


def _set_value(path: str, name: str, value: str) -> None:
    if winreg is None:
        raise OSError("Windows registry access is unavailable.")
    with winreg.CreateKeyEx(winreg.HKEY_CURRENT_USER, path, 0, winreg.KEY_SET_VALUE) as key:
        winreg.SetValueEx(key, name, 0, winreg.REG_SZ, value)


def _read_value(path: str, name: str) -> str | None:
    if winreg is None:
        return None
    try:
        with winreg.OpenKey(winreg.HKEY_CURRENT_USER, path, 0, winreg.KEY_READ) as key:
            value, _ = winreg.QueryValueEx(key, name)
            return str(value) if value else None
    except OSError:
        return None


def _read_default(path: str) -> str | None:
    return _read_value(path, "")


def _delete_tree(path: str) -> None:
    if winreg is None:
        return
    try:
        winreg.DeleteKeyEx(winreg.HKEY_CURRENT_USER, path, access=winreg.KEY_WRITE)
        return
    except OSError:
        pass
    try:
        with winreg.OpenKey(winreg.HKEY_CURRENT_USER, path, 0, winreg.KEY_READ | winreg.KEY_WRITE) as key:
            while True:
                try:
                    child = winreg.EnumKey(key, 0)
                except OSError:
                    break
                _delete_tree(path + "\\" + child)
        winreg.DeleteKey(winreg.HKEY_CURRENT_USER, path)
    except OSError:
        pass

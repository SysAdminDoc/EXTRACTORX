"""Post-extraction cleanup, post-actions, and external processors."""

from __future__ import annotations

import ctypes
import os
import shutil
import subprocess
from pathlib import Path
from typing import Callable
from uuid import uuid4

from .archive import archive_name


LogCallback = Callable[[str, str], None]


def propagate_motw(archive: Path, output: Path, log: LogCallback) -> None:
    """Copy the Zone.Identifier ADS from *archive* to all files under *output*.

    When a user downloads an archive from the internet, Windows tags it with
    a Zone.Identifier alternate data stream. Without propagation, extracted
    files lose their MOTW and bypass SmartScreen protection.
    """
    if os.name != "nt":
        return
    zone_stream = str(archive) + ":Zone.Identifier"
    try:
        zone_data = Path(zone_stream).read_text(encoding="utf-8", errors="replace")
    except OSError:
        return
    if not zone_data.strip():
        return
    propagated = 0
    for path in output.rglob("*"):
        if not path.is_file():
            continue
        try:
            target_stream = str(path) + ":Zone.Identifier"
            Path(target_stream).write_text(zone_data, encoding="utf-8")
            propagated += 1
        except OSError:
            continue
    if propagated:
        log(f"Propagated MOTW to {propagated} extracted file(s).", "info")


def cleanup_success_output(output: Path, archive: Path, config: dict, log: LogCallback) -> Path:
    current = output
    mode = str(config.get("SmartExtract", "Auto"))
    if mode == "AlwaysWrap":
        pass  # output already wraps contents; nothing to reshape.
    elif mode == "NeverWrap":
        flattened = _flatten_single_child(current, log)
        if flattened:
            current = flattened
    else:  # Auto: legacy duplicate-folder removal + tarbomb wrap guard.
        if bool(config.get("RemoveDuplicateFolder", True)):
            flattened = _remove_duplicate_root_folder(current, archive, log)
            if flattened:
                current = flattened
    if bool(config.get("RenameSingleFile", False)):
        _rename_single_file(current, archive, log)
    return current


def _flatten_single_child(output: Path, log: LogCallback) -> Path | None:
    """Pull a lone wrapping folder's contents up to ``output`` itself.

    Unlike ``_remove_duplicate_root_folder`` this ignores the archive name --
    it simply unwraps any single-subdirectory layout (``NeverWrap`` mode).
    """
    if not output.is_dir():
        return None
    try:
        children = list(output.iterdir())
    except OSError:
        return None
    if len(children) != 1 or not children[0].is_dir():
        return None
    wrapper = children[0]
    temp = output.with_name(f"{output.name}_flat_{uuid4().hex[:8]}")
    try:
        output.rename(temp)
        (temp / wrapper.name).rename(output)
        shutil.rmtree(temp, ignore_errors=True)
        log("Smart Extract: flattened wrapping folder.", "info")
        return output
    except OSError as exc:
        log(f"Smart Extract flatten failed: {exc}", "warning")
        try:
            if temp.exists() and not output.exists():
                temp.rename(output)
        except OSError:
            pass
    return None


def cleanup_failed_output(output: Path | None, config: dict, log: LogCallback) -> None:
    if not output or not bool(config.get("DeleteBrokenFiles", False)):
        return
    try:
        if output.is_dir():
            shutil.rmtree(output)
        elif output.exists():
            output.unlink()
        else:
            return
        log(f"Deleted incomplete output: {output}", "warning")
    except OSError as exc:
        log(f"Could not delete incomplete output: {exc}", "warning")


def apply_post_action(archive: Path, config: dict, log: LogCallback) -> None:
    if not archive.exists():
        return
    action = "Delete" if bool(config.get("DeleteAfterExtract", False)) else str(config.get("PostAction", "None"))
    try:
        if action == "Recycle":
            recycle_path(archive)
            log(f"Recycled source archive: {archive.name}", "info")
        elif action == "MoveToFolder":
            raw_destination = str(config.get("PostActionFolder", "") or "").strip()
            if not raw_destination:
                log("Post-action MoveToFolder skipped: no destination folder configured.", "warning")
                return
            destination = Path(raw_destination).expanduser()
            destination.mkdir(parents=True, exist_ok=True)
            target = _unique_path(destination / archive.name)
            shutil.move(str(archive), str(target))
            log(f"Moved source archive: {target}", "info")
        elif action == "Delete":
            archive.unlink()
            log(f"Deleted source archive: {archive.name}", "info")
    except OSError as exc:
        log(f"Post-action failed for {archive.name}: {exc}", "warning")


def open_destination(output: Path, log: LogCallback) -> None:
    try:
        if os.name == "nt":
            os.startfile(output)  # type: ignore[attr-defined]
        else:
            subprocess.Popen(["xdg-open", str(output)])
    except OSError as exc:
        log(f"Could not open destination: {exc}", "warning")


def run_external_processors(config: dict, archive: Path, output: Path, log: LogCallback) -> None:
    archive_extension = archive.suffix.lstrip(".").lower()
    for entry in config.get("ExternalProcessors", []):
        extension = str(entry.get("Extension", "")).strip().lstrip(".").lower()
        template = str(entry.get("Command", "")).strip()
        if not extension or not template or extension != archive_extension:
            continue
        command = expand_processor_command(template, archive=archive, output=output)
        try:
            result = subprocess.run(
                command,
                shell=True,
                cwd=str(output if output.is_dir() else output.parent),
                capture_output=True,
                text=True,
                encoding="utf-8",
                errors="replace",
                creationflags=subprocess.CREATE_NO_WINDOW if hasattr(subprocess, "CREATE_NO_WINDOW") else 0,
                timeout=3600,
                check=False,
            )
            if result.returncode == 0:
                log(f"External processor completed for {archive.name}", "success")
            else:
                tail = (result.stderr or result.stdout or "").strip().splitlines()[-1:] or [f"exit code {result.returncode}"]
                log(f"External processor failed for {archive.name}: {tail[0]}", "warning")
        except (OSError, subprocess.TimeoutExpired) as exc:
            log(f"External processor failed for {archive.name}: {exc}", "warning")


def expand_processor_command(template: str, archive: Path, output: Path) -> str:
    """Substitute archive/destination tokens into an external-processor command.

    Values are wrapped in double quotes and any embedded double quotes are
    escaped so paths with spaces or ``&`` / ``|`` meta-characters survive the
    trip through ``cmd.exe`` with ``shell=True``. User-supplied quotes around
    a token are absorbed so ``"{ArchivePath}"`` and ``{ArchivePath}`` behave
    identically.
    """

    def quote(value: str) -> str:
        return '"' + value.replace('"', '""') + '"'

    replacements = {
        "{ArchivePath}": quote(str(archive)),
        "{Destination}": quote(str(output)),
        "{Output}": quote(str(output)),
        "{ArchiveName}": quote(archive_name(archive)),
    }
    result = template
    for token, value in replacements.items():
        result = result.replace(f'"{token}"', value).replace(token, value)
    return result


def recycle_path(path: Path) -> None:
    """Send ``path`` to the Windows Recycle Bin (falls back to plain delete elsewhere).

    ``SHFileOperationW`` requires an absolute path terminated by two NUL
    characters. ``c_wchar_p`` bound to a Python string containing an embedded
    NUL can be truncated by ctypes, so we build the buffer manually with
    :func:`ctypes.create_unicode_buffer` and pass its address.
    """
    if os.name != "nt":
        if path.is_file():
            path.unlink()
        else:
            shutil.rmtree(path)
        return

    try:
        absolute = str(path.resolve(strict=False))
    except OSError as exc:  # pragma: no cover - only hits on transient FS errors
        raise OSError(f"Could not resolve path for recycle: {exc}") from exc

    class SHFILEOPSTRUCTW(ctypes.Structure):
        _fields_ = [
            ("hwnd", ctypes.c_void_p),
            ("wFunc", ctypes.c_uint),
            ("pFrom", ctypes.c_wchar_p),
            ("pTo", ctypes.c_wchar_p),
            ("fFlags", ctypes.c_ushort),
            ("fAnyOperationsAborted", ctypes.c_bool),
            ("hNameMappings", ctypes.c_void_p),
            ("lpszProgressTitle", ctypes.c_wchar_p),
        ]

    buffer = ctypes.create_unicode_buffer(absolute + "\0", len(absolute) + 2)
    operation = SHFILEOPSTRUCTW()
    operation.wFunc = 3  # FO_DELETE
    operation.pFrom = ctypes.cast(buffer, ctypes.c_wchar_p)
    operation.fFlags = 0x0040 | 0x0010 | 0x0400  # allow undo, no confirmation, no error UI
    result = ctypes.windll.shell32.SHFileOperationW(ctypes.byref(operation))
    if result != 0 or operation.fAnyOperationsAborted:
        raise OSError(f"Recycle failed with code {result}")


def _remove_duplicate_root_folder(output: Path, archive: Path, log: LogCallback) -> Path | None:
    if not output.is_dir():
        return None
    expected = archive_name(archive)
    try:
        children = list(output.iterdir())
    except OSError:
        return None
    if len(children) != 1 or not children[0].is_dir() or children[0].name.lower() != expected.lower():
        return None

    temp = output.with_name(f"{output.name}_flat_{uuid4().hex[:8]}")
    try:
        output.rename(temp)
        (temp / children[0].name).rename(output)
        shutil.rmtree(temp, ignore_errors=True)
        log("Removed duplicate archive-name folder.", "info")
        return output
    except OSError as exc:
        log(f"Could not remove duplicate folder: {exc}", "warning")
        try:
            if temp.exists() and not output.exists():
                temp.rename(output)
        except OSError:
            pass
    return None


def _rename_single_file(output: Path, archive: Path, log: LogCallback) -> None:
    if not output.is_dir():
        return
    try:
        files = [path for path in output.iterdir() if path.is_file()]
        dirs = [path for path in output.iterdir() if path.is_dir()]
    except OSError:
        return
    if len(files) != 1 or dirs:
        return
    source = files[0]
    target = _unique_path(source.with_name(archive_name(archive) + source.suffix))
    if source == target:
        return
    try:
        source.rename(target)
        log(f"Renamed single file: {source.name} -> {target.name}", "info")
    except OSError as exc:
        log(f"Could not rename single file: {exc}", "warning")


def _unique_path(path: Path) -> Path:
    if not path.exists():
        return path
    counter = 1
    while True:
        candidate = path.with_name(f"{path.stem} ({counter}){path.suffix}")
        if not candidate.exists():
            return candidate
        counter += 1

"""Configuration loading, normalization, and persistence."""

from __future__ import annotations

import json
import logging
import os
from copy import deepcopy
from pathlib import Path
from typing import Any
from uuid import uuid4


log = logging.getLogger("extractorx.config")


APP_NAME = "ExtractorX"
CONFIG_FILE = "config.json"
PORTABLE_FLAG = "portable.flag"

THEMES = ("Midnight", "Graphite", "Ocean", "White", "HighContrast")
OVERWRITE_MODES = ("Always", "Never", "Rename")
POST_ACTIONS = ("None", "Recycle", "MoveToFolder", "Delete")
DRAG_DROP_FILTERS = ("None", "Inclusion", "Exclusion")
THREAD_PRIORITIES = ("Low", "BelowNormal", "Normal", "AboveNormal", "High")
SMART_EXTRACT_MODES = ("Auto", "AlwaysWrap", "NeverWrap")
FILENAME_ENCODINGS = ("Auto", "UTF-8", "cp437", "cp932", "cp936", "cp949", "cp950", "cp1251", "cp1252")

DEFAULT_CONFIG: dict[str, Any] = {
    "OutputPath": r"{ArchiveFolder}\{ArchiveName}",
    "OverwriteMode": "Always",
    "PostAction": "None",
    "PostActionFolder": "",
    "DeleteAfterExtract": False,
    "OpenDestAfterExtract": False,
    "NestedExtraction": True,
    "NestedMaxDepth": 5,
    "NestedApplyPostAction": False,
    "RemoveDuplicateFolder": True,
    "RenameSingleFile": False,
    "DeleteBrokenFiles": False,
    "CloseOnComplete": False,
    "CloseOnCompleteAlways": False,
    "ClearListOnComplete": False,
    "AlwaysOnTop": False,
    "MinimizeToTray": True,
    "LogHistory": True,
    "AutoSwitchToHistory": True,
    "DeepArchiveDetection": True,
    "Theme": "Midnight",
    "AutoExtractOnDrop": False,
    "DragDropFilterType": "None",
    "DragDropFilterMask": "",
    "UsePasswordList": True,
    "PromptOnExhaustion": False,
    "AssumeOnePassword": True,
    "PasswordTimeout": 45,
    "FileExclusions": "Thumbs.db;desktop.ini;.DS_Store",
    "WatchFolders": [],
    "WatchFolderRules": [],
    "WatchAutoExtract": True,
    "CtxEnabled": False,
    "CtxGrouped": True,
    "CtxExtractHere": True,
    "CtxExtractToFolder": True,
    "CtxEnqueue": True,
    "CtxSearchArchives": True,
    "FileAssociations": [],
    "ThreadPriority": "Normal",
    "SoundsEnabled": True,
    "ExternalProcessors": [],
    "WindowWidth": 1100,
    "WindowHeight": 750,
    "WindowLeft": -1,
    "WindowTop": -1,
    "SmartExtract": "Auto",
    "FilenameEncoding": "Auto",
    "IncludeMasks": "",
    "SkipAfterFailedPasswords": 0,
    "UsePasswordSidecars": True,
    "HashModePasswordProbe": True,
    "WordlistGeneration": False,
    "WordlistMaxAttempts": 500,
    "PreExtractCommand": "",
    "PostExtractCommand": "",
    "OnFailureCommand": "",
    "SevenZipOverride": "",
    "MaxParallelExtractions": 1,
    "MaxDecompressionRatio": 1000,
    "PropagateMotw": True,
    "RetryCount": 0,
    "RetryDelaySeconds": 30,
    "SecureDelete": False,
    "DiskSpaceCheck": True,
    "WebhookUrl": "",
    "PasswordRules": [],
    "HandlerAllowlist": [],
    "Bookmarks": [],
}


_PORTABLE_ROOT_CACHE: tuple[bool, Path | None] = (False, None)


def _portable_root() -> Path | None:
    """Locate the portable install root if present.

    Cached after the first call because ``app_data_dir`` is called from many
    places on every config/log write. The cache keys off ``sys.argv[0]`` and
    the module location, both of which never change during a run.
    """
    global _PORTABLE_ROOT_CACHE
    if _PORTABLE_ROOT_CACHE[0]:
        return _PORTABLE_ROOT_CACHE[1]
    import sys

    candidates: list[Path] = []
    if getattr(sys, "frozen", False):
        try:
            candidates.append(Path(sys.executable).resolve().parent)
        except OSError:
            pass
    try:
        if sys.argv and sys.argv[0]:
            script = Path(sys.argv[0]).expanduser()
            if script.exists():
                candidates.append(script.resolve().parent)
    except OSError:
        pass
    try:
        candidates.append(Path(__file__).resolve().parents[1])
    except OSError:
        pass
    found: Path | None = None
    for candidate in candidates:
        try:
            if (candidate / PORTABLE_FLAG).exists():
                found = candidate
                break
        except OSError:
            continue
    _PORTABLE_ROOT_CACHE = (True, found)
    return found


def _reset_portable_cache() -> None:
    """Test hook: clear the portable-root cache."""
    global _PORTABLE_ROOT_CACHE
    _PORTABLE_ROOT_CACHE = (False, None)


def app_data_dir() -> Path:
    portable = _portable_root()
    if portable:
        return portable / f"{APP_NAME}.data"
    root = os.environ.get("APPDATA")
    if root:
        return Path(root) / APP_NAME
    return Path.home() / ".extractorx"


def is_portable() -> bool:
    return _portable_root() is not None


def config_path() -> Path:
    return app_data_dir() / CONFIG_FILE


def ensure_app_dirs() -> None:
    app_data_dir().mkdir(parents=True, exist_ok=True)
    (app_data_dir() / "logs").mkdir(parents=True, exist_ok=True)


def _as_list(value: Any) -> list[Any]:
    if value is None:
        return []
    if isinstance(value, list):
        return value
    if isinstance(value, tuple):
        return list(value)
    if isinstance(value, str):
        return [value] if value.strip() else []
    return [value]


def _as_bool(value: Any) -> bool:
    if isinstance(value, bool):
        return value
    if isinstance(value, str):
        return value.strip().lower() in {"1", "true", "yes", "on"}
    return bool(value)


def _clamp_int(value: Any, default: int, minimum: int, maximum: int) -> int:
    try:
        parsed = int(value)
    except (TypeError, ValueError):
        parsed = default
    return max(minimum, min(maximum, parsed))


def normalize_config(raw: dict[str, Any] | None) -> dict[str, Any]:
    config = deepcopy(DEFAULT_CONFIG)
    if raw:
        for key in DEFAULT_CONFIG:
            if key in raw and raw[key] is not None:
                config[key] = raw[key]

    bool_keys = (
        "DeleteAfterExtract",
        "OpenDestAfterExtract",
        "NestedExtraction",
        "NestedApplyPostAction",
        "RemoveDuplicateFolder",
        "RenameSingleFile",
        "DeleteBrokenFiles",
        "CloseOnComplete",
        "CloseOnCompleteAlways",
        "ClearListOnComplete",
        "AlwaysOnTop",
        "MinimizeToTray",
        "LogHistory",
        "AutoSwitchToHistory",
        "DeepArchiveDetection",
        "AutoExtractOnDrop",
        "UsePasswordList",
        "PromptOnExhaustion",
        "AssumeOnePassword",
        "WatchAutoExtract",
        "CtxEnabled",
        "CtxGrouped",
        "CtxExtractHere",
        "CtxExtractToFolder",
        "CtxEnqueue",
        "CtxSearchArchives",
        "SoundsEnabled",
        "UsePasswordSidecars",
        "HashModePasswordProbe",
        "WordlistGeneration",
        "PropagateMotw",
        "SecureDelete",
    )
    for key in bool_keys:
        config[key] = _as_bool(config[key])

    if config["Theme"] not in THEMES:
        config["Theme"] = "Midnight"
    if config["OverwriteMode"] not in OVERWRITE_MODES:
        config["OverwriteMode"] = "Always"
    if config["PostAction"] not in POST_ACTIONS:
        config["PostAction"] = "None"
    if config["DragDropFilterType"] not in DRAG_DROP_FILTERS:
        config["DragDropFilterType"] = "None"
    if config["ThreadPriority"] not in THREAD_PRIORITIES:
        config["ThreadPriority"] = "Normal"
    if config["SmartExtract"] not in SMART_EXTRACT_MODES:
        config["SmartExtract"] = "Auto"
    if config["FilenameEncoding"] not in FILENAME_ENCODINGS:
        config["FilenameEncoding"] = "Auto"

    config["NestedMaxDepth"] = _clamp_int(config["NestedMaxDepth"], 5, 1, 50)
    config["PasswordTimeout"] = _clamp_int(config["PasswordTimeout"], 45, 5, 600)
    config["WindowWidth"] = _clamp_int(config["WindowWidth"], 1100, 850, 5000)
    config["WindowHeight"] = _clamp_int(config["WindowHeight"], 750, 500, 4000)
    config["WindowLeft"] = _clamp_int(config["WindowLeft"], -1, -1, 100000)
    config["WindowTop"] = _clamp_int(config["WindowTop"], -1, -1, 100000)
    config["MaxParallelExtractions"] = _clamp_int(config["MaxParallelExtractions"], 1, 1, 8)
    config["SkipAfterFailedPasswords"] = _clamp_int(config["SkipAfterFailedPasswords"], 0, 0, 9999)
    config["WordlistMaxAttempts"] = _clamp_int(config["WordlistMaxAttempts"], 500, 10, 10000)
    config["MaxDecompressionRatio"] = _clamp_int(config["MaxDecompressionRatio"], 1000, 0, 100000)
    config["RetryCount"] = _clamp_int(config["RetryCount"], 0, 0, 10)
    config["RetryDelaySeconds"] = _clamp_int(config["RetryDelaySeconds"], 30, 5, 600)

    config["OutputPath"] = str(config["OutputPath"] or DEFAULT_CONFIG["OutputPath"])
    config["PostActionFolder"] = str(config["PostActionFolder"] or "")
    config["DragDropFilterMask"] = str(config["DragDropFilterMask"] or "")
    config["FileExclusions"] = str(config["FileExclusions"] or "")
    config["IncludeMasks"] = str(config["IncludeMasks"] or "")
    config["PreExtractCommand"] = str(config["PreExtractCommand"] or "")
    config["PostExtractCommand"] = str(config["PostExtractCommand"] or "")
    config["OnFailureCommand"] = str(config["OnFailureCommand"] or "")
    config["SevenZipOverride"] = str(config["SevenZipOverride"] or "")

    config["WatchFolders"] = [str(item) for item in _as_list(config["WatchFolders"]) if str(item).strip()]
    watch_rules: list[dict[str, str]] = []
    for entry in _as_list(config["WatchFolderRules"]):
        if isinstance(entry, dict):
            folder = str(entry.get("Folder", "")).strip()
            if folder:
                watch_rules.append({
                    "Folder": folder,
                    "OutputPath": str(entry.get("OutputPath", "")).strip(),
                    "PostAction": str(entry.get("PostAction", "")).strip(),
                })
    config["WatchFolderRules"] = watch_rules
    config["FileAssociations"] = [str(item) for item in _as_list(config["FileAssociations"]) if str(item).strip()]
    pw_rules: list[dict[str, object]] = []
    for entry in _as_list(config["PasswordRules"]):
        if isinstance(entry, dict):
            pattern = str(entry.get("Pattern", "")).strip()
            passwords_val = entry.get("Passwords", [])
            passwords_list = [str(p) for p in _as_list(passwords_val) if str(p).strip()]
            if pattern and passwords_list:
                pw_rules.append({"Pattern": pattern, "Passwords": passwords_list})
    config["PasswordRules"] = pw_rules

    config["HandlerAllowlist"] = [
        str(item).strip().lower().lstrip(".")
        for item in _as_list(config["HandlerAllowlist"])
        if str(item).strip()
    ]
    bookmarks: list[dict[str, str]] = []
    for entry in _as_list(config["Bookmarks"]):
        if isinstance(entry, dict):
            label = str(entry.get("Label", "")).strip()
            path = str(entry.get("Path", "")).strip()
            if label and path:
                bookmarks.append({"Label": label, "Path": path})
        elif isinstance(entry, str) and entry.strip():
            bookmarks.append({"Label": entry.strip(), "Path": entry.strip()})
    config["Bookmarks"] = bookmarks

    processors: list[dict[str, str]] = []
    for entry in _as_list(config["ExternalProcessors"]):
        if not isinstance(entry, dict):
            continue
        extension = str(entry.get("Extension", "")).strip().lstrip(".")
        command = str(entry.get("Command", "")).strip()
        if extension and command:
            processors.append({"Extension": extension, "Command": command})
    config["ExternalProcessors"] = processors
    return config


def _backup_corrupt_config(path: Path) -> None:
    """Rename a broken config file out of the way so a fresh one can be written."""
    if not path.exists():
        return
    try:
        backup = path.with_suffix(f".corrupt-{uuid4().hex[:8]}.json")
        path.rename(backup)
        log.warning("Moved unreadable config to %s and restored defaults.", backup)
    except OSError as exc:
        log.warning("Could not back up unreadable config %s: %s", path, exc)


def load_config() -> dict[str, Any]:
    ensure_app_dirs()
    path = config_path()
    if not path.exists():
        return normalize_config(None)
    try:
        raw = path.read_text(encoding="utf-8-sig")
    except OSError as exc:
        log.warning("Could not read config %s: %s -- using defaults.", path, exc)
        return normalize_config(None)
    try:
        parsed = json.loads(raw) if raw.strip() else None
    except json.JSONDecodeError as exc:
        log.warning("Config %s is not valid JSON (%s) -- using defaults.", path, exc)
        _backup_corrupt_config(path)
        return normalize_config(None)
    if parsed is not None and not isinstance(parsed, dict):
        log.warning("Config %s was not a JSON object -- using defaults.", path)
        _backup_corrupt_config(path)
        return normalize_config(None)
    return normalize_config(parsed)


def _atomic_write(path: Path, data: str, *, encoding: str = "utf-8") -> None:
    """Write ``data`` to ``path`` atomically.

    Writes to a sibling temp file, flushes + fsyncs, then swaps it into place
    with :func:`os.replace`. Partial writes never leave ``path`` corrupted --
    either the old content survives or the new content is fully present.
    """
    path.parent.mkdir(parents=True, exist_ok=True)
    temp = path.with_name(f"{path.name}.{uuid4().hex[:8]}.tmp")
    try:
        with temp.open("w", encoding=encoding, newline="\n") as handle:
            handle.write(data)
            handle.flush()
            try:
                os.fsync(handle.fileno())
            except OSError:
                pass
        os.replace(temp, path)
    finally:
        try:
            if temp.exists():
                temp.unlink()
        except OSError:
            pass


def save_config(config: dict[str, Any]) -> dict[str, Any]:
    ensure_app_dirs()
    normalized = normalize_config(config)
    _atomic_write(config_path(), json.dumps(normalized, indent=2))
    return normalized

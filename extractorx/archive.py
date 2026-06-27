"""Archive detection and path helpers."""

from __future__ import annotations

import os
import re
import zipfile
from uuid import uuid4
from datetime import datetime
from pathlib import Path


ARCHIVE_EXTENSIONS = {
    ".zip",
    ".7z",
    ".rar",
    ".tar",
    ".gz",
    ".gzip",
    ".tgz",
    ".bz2",
    ".bzip2",
    ".tbz2",
    ".tbz",
    ".xz",
    ".txz",
    ".lzma",
    ".tlz",
    ".lz",
    ".zst",
    ".zstd",
    ".z",
    ".iso",
    ".cab",
    ".arj",
    ".lzh",
    ".lha",
    ".wim",
    ".001",
    ".cpio",
    ".rpm",
    ".deb",
}

MAGIC_BYTES = (
    bytes.fromhex("504B0304"),
    bytes.fromhex("504B0506"),
    bytes.fromhex("504B0708"),
    bytes.fromhex("377ABCAF271C"),
    bytes.fromhex("526172211A0700"),
    bytes.fromhex("526172211A07"),
    bytes.fromhex("1F8B"),
    bytes.fromhex("425A68"),
    bytes.fromhex("FD377A585A00"),
    bytes.fromhex("4D534346"),
)


def is_archive_path(path: Path | str) -> bool:
    return Path(path).suffix.lower() in ARCHIVE_EXTENSIONS


def has_archive_magic(path: Path | str) -> bool:
    file_path = Path(path)
    try:
        file_size = file_path.stat().st_size
        with file_path.open("rb") as handle:
            head = handle.read(16)
            if any(head.startswith(signature) for signature in MAGIC_BYTES):
                return True
            if file_size > 262:
                handle.seek(257)
                if handle.read(5) == b"ustar":
                    return True
            if file_size > 32773:
                handle.seek(32769)
                if handle.read(5) == b"CD001":
                    return True
    except OSError:
        return False
    return False


def is_supported_archive(path: Path | str, deep_detection: bool = True) -> bool:
    candidate = Path(path)
    if not candidate.is_file():
        return False
    if is_archive_path(candidate):
        return True
    return deep_detection and has_archive_magic(candidate)


def is_non_first_volume(path: Path | str) -> bool:
    name = Path(path).name
    part = re.search(r"\.part(\d+)\.rar$", name, re.IGNORECASE)
    if part and int(part.group(1)) > 1:
        return True
    if re.search(r"\.[rs]\d{2}$", name, re.IGNORECASE):
        return True
    split = re.search(r"\.\w+\.(\d{3})$", name, re.IGNORECASE)
    return bool(split and split.group(1) != "001")


def archive_name(path: Path | str) -> str:
    candidate = Path(path)
    name = candidate.name
    for suffix in (".tar.gz", ".tar.bz2", ".tar.xz", ".tar.zst", ".tgz", ".tbz2", ".txz"):
        if name.lower().endswith(suffix):
            return name[: -len(suffix)]
    return candidate.stem


def resolve_output_path(template: str, archive: Path | str) -> Path:
    archive_path = Path(archive).expanduser()
    if not archive_path.is_absolute():
        archive_path = Path(os.path.abspath(archive_path))
    now = datetime.now()
    resolved = template or r"{ArchiveFolder}\{ArchiveName}"
    name = archive_name(archive_path)
    replacements = {
        "{ArchiveFolder}": str(archive_path.parent),
        "{ArchiveName}": name,
        "{ArchiveExtension}": archive_path.suffix.lstrip("."),
        "{ArchiveExt}": archive_path.suffix.lstrip("."),
        "{ArchiveFileName}": archive_path.name,
        "{ArchiveFolderName}": archive_path.parent.name,
        "{ArchivePath}": str(archive_path),
        "{Desktop}": str(Path.home() / "Desktop"),
        "{UserProfile}": str(Path.home()),
        "{Windows}": os.environ.get("windir", r"C:\Windows"),
        "{Program Files}": os.environ.get("ProgramFiles", r"C:\Program Files"),
        "{ProgramFiles}": os.environ.get("ProgramFiles", r"C:\Program Files"),
        "{Guid}": uuid4().hex[:8],
        "{Date}": now.strftime("%Y%m%d"),
        "{Time}": now.strftime("%H%M%S"),
    }
    for token, value in replacements.items():
        resolved = resolved.replace(token, value)

    def replace_env(match: re.Match[str]) -> str:
        return os.environ.get(match.group(1), "")

    resolved = re.sub(r"\{Env:([^}]+)\}", replace_env, resolved, flags=re.IGNORECASE)
    if "{ArchiveNameUnique}" in resolved:
        base = resolved.replace("{ArchiveNameUnique}", name)
        if Path(base).exists():
            for counter in range(1, 10001):
                candidate = resolved.replace("{ArchiveNameUnique}", f"{name} ({counter})")
                if not Path(candidate).exists():
                    resolved = candidate
                    break
            else:
                resolved = resolved.replace("{ArchiveNameUnique}", f"{name} ({uuid4().hex[:8]})")
        else:
            resolved = base
    return Path(resolved).expanduser()


_CODEPAGE_CANDIDATES = [
    ("cp932", "Japanese"),
    ("cp936", "Simplified Chinese"),
    ("cp949", "Korean"),
    ("cp950", "Traditional Chinese"),
    ("cp1251", "Cyrillic"),
    ("cp1252", "Western European"),
    ("cp437", "DOS/OEM"),
]


def detect_zip_codepage(path: Path) -> str | None:
    """Heuristically detect the filename encoding of a ZIP archive.

    Returns a codepage string (e.g. ``'cp932'``) if non-UTF-8 filenames are
    detected and a likely encoding is found, or ``None`` if the archive is
    UTF-8 or detection is inconclusive.
    """
    if not path.is_file() or path.suffix.lower() != ".zip":
        return None
    try:
        with zipfile.ZipFile(path, "r") as zf:
            raw_names: list[bytes] = []
            for info in zf.infolist():
                if info.flag_bits & 0x800:
                    continue
                try:
                    info.filename.encode("ascii")
                except UnicodeEncodeError:
                    pass
                else:
                    continue
                raw = info.filename.encode("cp437", errors="replace")
                raw_names.append(raw)
    except (zipfile.BadZipFile, OSError):
        return None
    if not raw_names:
        return None
    best_cp = None
    best_score = 0
    for cp, _label in _CODEPAGE_CANDIDATES:
        score = 0
        for raw in raw_names:
            try:
                decoded = raw.decode(cp)
                if all(c.isprintable() or c in (" ", "/", "\\") for c in decoded):
                    score += 1
            except (UnicodeDecodeError, LookupError):
                continue
        if score > best_score:
            best_score = score
            best_cp = cp
    if best_cp and best_score == len(raw_names):
        return best_cp
    return None


def validate_extraction_paths(output: Path) -> list[Path]:
    """Walk *output* and return any files that resolve outside it.

    Defense-in-depth against zip-slip / symlink path traversal. Call after
    7-Zip extraction completes to verify nothing escaped the output dir.
    """
    escaped: list[Path] = []
    try:
        canonical_root = output.resolve(strict=True)
    except OSError:
        return escaped
    norm_root = os.path.normcase(str(canonical_root))
    try:
        for path in output.rglob("*"):
            try:
                if path.is_symlink() or (os.name == "nt" and _is_reparse_point(path)):
                    escaped.append(path)
                    continue
                canonical = path.resolve(strict=True)
                norm_canonical = os.path.normcase(str(canonical))
                if not norm_canonical.startswith(norm_root + os.sep) and norm_canonical != norm_root:
                    escaped.append(path)
            except OSError:
                continue
    except OSError:
        pass
    return escaped


def _is_reparse_point(path: Path) -> bool:
    """Check if *path* is a Windows reparse point (junction, symlink, etc.)."""
    try:
        import ctypes
        attrs = ctypes.windll.kernel32.GetFileAttributesW(str(path))
        return attrs != -1 and bool(attrs & 0x0400)
    except (OSError, AttributeError):
        return False


def format_size(size: int) -> str:
    units = ("B", "KB", "MB", "GB", "TB")
    value = float(max(0, size))
    for unit in units:
        if value < 1024 or unit == units[-1]:
            return f"{value:.1f} {unit}" if unit != "B" else f"{int(value)} B"
        value /= 1024
    return f"{int(size)} B"

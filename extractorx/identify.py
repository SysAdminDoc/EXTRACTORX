"""File identification for "scan-only" mode.

Maps the magic-byte table plus extension heuristics to a human-readable
label + supported-by-7z flag, without unpacking anything. Used by the
``--identify`` CLI entrypoint and a Settings > Identify button.
"""

from __future__ import annotations

from dataclasses import dataclass
from pathlib import Path

from .archive import ARCHIVE_EXTENSIONS, has_archive_magic, is_archive_path


@dataclass(frozen=True)
class IdentifyResult:
    path: Path
    format: str
    supported: bool
    reason: str


_MAGIC_LABELS: list[tuple[bytes, str]] = [
    (bytes.fromhex("504B0304"), "ZIP"),
    (bytes.fromhex("504B0506"), "ZIP (empty)"),
    (bytes.fromhex("504B0708"), "ZIP (spanned)"),
    (bytes.fromhex("377ABCAF271C"), "7z"),
    (bytes.fromhex("526172211A0700"), "RAR 5"),
    (bytes.fromhex("526172211A07"), "RAR 4"),
    (bytes.fromhex("1F8B"), "gzip"),
    (bytes.fromhex("425A68"), "bzip2"),
    (bytes.fromhex("FD377A585A00"), "xz"),
    (bytes.fromhex("4D534346"), "Microsoft Cabinet"),
]


def identify(path: Path | str) -> IdentifyResult:
    candidate = Path(path)
    if not candidate.exists():
        return IdentifyResult(candidate, "missing", False, "File does not exist")
    if not candidate.is_file():
        return IdentifyResult(candidate, "directory", False, "Not a regular file")
    header = b""
    try:
        with candidate.open("rb") as handle:
            header = handle.read(16)
    except OSError as exc:
        return IdentifyResult(candidate, "unreadable", False, str(exc))
    for signature, label in _MAGIC_LABELS:
        if header.startswith(signature):
            return IdentifyResult(candidate, label, True, "Detected by magic bytes")
    try:
        with candidate.open("rb") as handle:
            size = candidate.stat().st_size
            if size > 262:
                handle.seek(257)
                if handle.read(5) == b"ustar":
                    return IdentifyResult(candidate, "tar", True, "ustar magic")
            if size > 32773:
                handle.seek(32769)
                if handle.read(5) == b"CD001":
                    return IdentifyResult(candidate, "ISO 9660", True, "ISO CD001 marker")
    except OSError:
        pass
    if is_archive_path(candidate):
        return IdentifyResult(
            candidate,
            f"{candidate.suffix.lstrip('.').lower()} (by extension)",
            True,
            "Extension is in the supported list",
        )
    if has_archive_magic(candidate):
        return IdentifyResult(candidate, "archive (magic)", True, "Secondary magic check passed")
    return IdentifyResult(
        candidate,
        "unknown",
        False,
        f"No matching signature; known extensions: {', '.join(sorted(ARCHIVE_EXTENSIONS))}",
    )

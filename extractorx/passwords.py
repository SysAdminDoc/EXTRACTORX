"""Password storage using Windows DPAPI when available."""

from __future__ import annotations

import base64
import ctypes
import json
import logging
import math
import os
import string
from ctypes import wintypes
from pathlib import Path
from uuid import uuid4

from .config import app_data_dir


log = logging.getLogger("extractorx.passwords")


def estimate_entropy_bits(password: str) -> float:
    """Rough Shannon-style entropy estimate in bits.

    Lighter than ``zxcvbn`` -- good enough to drive a "weak/fair/strong" label
    next to the password input without pulling in an extra dependency.
    """
    if not password:
        return 0.0
    classes = 0
    if any(char in string.ascii_lowercase for char in password):
        classes += 26
    if any(char in string.ascii_uppercase for char in password):
        classes += 26
    if any(char in string.digits for char in password):
        classes += 10
    if any(char in string.punctuation for char in password):
        classes += len(string.punctuation)
    if any(char > "~" for char in password):
        classes += 128
    classes = max(classes, len(set(password)))
    if classes <= 1:
        return 0.0
    return len(password) * math.log2(classes)


def classify_entropy(password: str) -> tuple[str, float]:
    bits = estimate_entropy_bits(password)
    if bits < 28:
        label = "weak"
    elif bits < 60:
        label = "fair"
    else:
        label = "strong"
    return label, bits


_LEET_MAP = str.maketrans({"a": "4", "e": "3", "i": "1", "o": "0", "s": "5", "t": "7"})


def generate_wordlist(
    passwords: list[str],
    *,
    case_variants: bool = True,
    leet_variants: bool = True,
    date_suffixes: bool = True,
    max_total: int = 500,
) -> list[str]:
    """Expand *passwords* with case, leet-speak, and date-suffix permutations.

    Returns at most *max_total* unique candidates. The original passwords are
    always included first so the caller can feed the result straight into the
    password-attempt pipeline.
    """
    from datetime import datetime

    seen: set[str] = set()
    result: list[str] = []

    def add(candidate: str) -> bool:
        if candidate in seen or not candidate:
            return True
        if len(result) >= max_total:
            return False
        seen.add(candidate)
        result.append(candidate)
        return True

    for password in passwords:
        if not add(password):
            return result

    for password in passwords:
        if case_variants:
            if not add(password.lower()):
                return result
            if not add(password.upper()):
                return result
            if not add(password.capitalize()):
                return result
            if not add(password.swapcase()):
                return result
        if leet_variants:
            leet = password.lower().translate(_LEET_MAP)
            if not add(leet):
                return result
            if not add(leet.capitalize()):
                return result
        if date_suffixes:
            now = datetime.now()
            for suffix in (
                str(now.year),
                now.strftime("%m%Y"),
                now.strftime("%Y%m"),
                str(now.year - 1),
                "!",
                "1",
                "123",
            ):
                if not add(password + suffix):
                    return result
    return result


_DPAPI_ENTROPY = b"ExtractorX-v2-password-store"


class DATA_BLOB(ctypes.Structure):
    _fields_ = [("cbData", wintypes.DWORD), ("pbData", ctypes.POINTER(ctypes.c_byte))]


class PasswordStore:
    def __init__(self, path: Path | None = None) -> None:
        self.path = path or (app_data_dir() / "passwords.dat")
        self._migrate_legacy_filename()

    def _migrate_legacy_filename(self) -> None:
        """Rename the 2.1 port's `passwords.py.dat` to `passwords.dat` if present.

        Only applies to the default app-data location so user-specified paths
        are left untouched.
        """
        if self.path.name != "passwords.dat":
            return
        legacy = self.path.with_name("passwords.py.dat")
        if not self.path.exists() and legacy.exists():
            try:
                legacy.rename(self.path)
            except OSError:
                pass

    def load(self) -> list[str]:
        if not self.path.exists():
            return []
        try:
            raw = self.path.read_bytes()
        except OSError as exc:
            log.warning("Could not read password store %s: %s", self.path, exc)
            return []
        try:
            text = _unprotect(raw)
            parsed = json.loads(text)
        except (OSError, json.JSONDecodeError, UnicodeDecodeError, ValueError) as exc:
            log.warning("Password store %s is unreadable (%s); treating as empty.", self.path, exc)
            return []
        except Exception as exc:  # DPAPI failures surface as OSError/WinError subclasses
            log.warning("Password store %s could not be decrypted (%s); treating as empty.", self.path, exc)
            return []
        if not isinstance(parsed, list):
            return []
        passwords = [str(item) for item in parsed if str(item)]
        if raw.startswith(b"dpapi:") and os.name == "nt":
            try:
                self.save(passwords)
                log.info("Migrated password store to entropy-protected format.")
            except Exception:
                pass
        return passwords

    def save(self, passwords: list[str]) -> None:
        self.path.parent.mkdir(parents=True, exist_ok=True)
        unique: list[str] = []
        for password in passwords:
            if password and password not in unique:
                unique.append(password)
        payload = json.dumps(unique).encode("utf-8")
        encrypted = _protect(payload)
        temp = self.path.with_name(f"{self.path.name}.{uuid4().hex[:8]}.tmp")
        try:
            with temp.open("wb") as handle:
                handle.write(encrypted)
                handle.flush()
                try:
                    os.fsync(handle.fileno())
                except OSError:
                    pass
            os.replace(temp, self.path)
        finally:
            try:
                if temp.exists():
                    temp.unlink()
            except OSError:
                pass


def _protect(data: bytes) -> bytes:
    if os.name != "nt":
        return b"plain:" + base64.b64encode(data)
    return b"dpapi2:" + _crypt_protect(data, _DPAPI_ENTROPY)


def _unprotect(data: bytes) -> str:
    if data.startswith(b"dpapi2:") and os.name == "nt":
        return _crypt_unprotect(data[7:], _DPAPI_ENTROPY).decode("utf-8")
    if data.startswith(b"dpapi:") and os.name == "nt":
        return _crypt_unprotect(data[6:]).decode("utf-8")
    if data.startswith(b"plain:"):
        return base64.b64decode(data[6:]).decode("utf-8")
    return data.decode("utf-8")


def _blob_from_bytes(data: bytes) -> tuple[DATA_BLOB, ctypes.Array[ctypes.c_byte]]:
    buffer = ctypes.create_string_buffer(data)
    blob = DATA_BLOB(len(data), ctypes.cast(buffer, ctypes.POINTER(ctypes.c_byte)))
    return blob, buffer


def _crypt_protect(data: bytes, entropy: bytes | None = None) -> bytes:
    crypt32 = ctypes.windll.crypt32
    kernel32 = ctypes.windll.kernel32
    kernel32.LocalFree.argtypes = [wintypes.HLOCAL]
    kernel32.LocalFree.restype = wintypes.HLOCAL
    in_blob, in_buffer = _blob_from_bytes(data)
    if entropy:
        ent_blob, ent_buffer = _blob_from_bytes(entropy)
        ent_arg = ctypes.byref(ent_blob)
    else:
        ent_arg = None
        ent_buffer = None
    out_blob = DATA_BLOB()
    try:
        if not crypt32.CryptProtectData(ctypes.byref(in_blob), None, ent_arg, None, None, 0, ctypes.byref(out_blob)):
            raise ctypes.WinError()
        return ctypes.string_at(out_blob.pbData, out_blob.cbData)
    finally:
        _ = in_buffer
        _ = ent_buffer
        if out_blob.pbData:
            kernel32.LocalFree(ctypes.cast(out_blob.pbData, wintypes.HLOCAL))


def _crypt_unprotect(data: bytes, entropy: bytes | None = None) -> bytes:
    crypt32 = ctypes.windll.crypt32
    kernel32 = ctypes.windll.kernel32
    kernel32.LocalFree.argtypes = [wintypes.HLOCAL]
    kernel32.LocalFree.restype = wintypes.HLOCAL
    in_blob, in_buffer = _blob_from_bytes(data)
    if entropy:
        ent_blob, ent_buffer = _blob_from_bytes(entropy)
        ent_arg = ctypes.byref(ent_blob)
    else:
        ent_arg = None
        ent_buffer = None
    out_blob = DATA_BLOB()
    try:
        if not crypt32.CryptUnprotectData(ctypes.byref(in_blob), None, ent_arg, None, None, 0, ctypes.byref(out_blob)):
            raise ctypes.WinError()
        return ctypes.string_at(out_blob.pbData, out_blob.cbData)
    finally:
        _ = in_buffer
        _ = ent_buffer
        if out_blob.pbData:
            kernel32.LocalFree(ctypes.cast(out_blob.pbData, wintypes.HLOCAL))

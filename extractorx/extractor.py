"""Extraction service backed by 7-Zip."""

from __future__ import annotations

import logging
import re as _re
import shutil
import subprocess
import ctypes
import os
from concurrent.futures import ThreadPoolExecutor, wait
from pathlib import Path
from queue import Queue
from threading import Event, Lock, Thread


log = logging.getLogger("extractorx.extractor")

from .archive import archive_name, detect_zip_codepage, is_non_first_volume, is_supported_archive, resolve_output_path, validate_extraction_paths
from .hooks import run_hook
from .models import OperationMessage, QueueItem, QueueStatus
from .plugins import run_plugins
from .postprocess import (
    apply_post_action,
    cleanup_failed_output,
    cleanup_success_output,
    open_destination,
    propagate_motw,
    run_external_processors,
)
from .passwords import generate_wordlist
from .scanner import scan_paths
from .sevenzip import build_sevenzip_command, check_7zip_version, download_7zip, MIN_SEVENZIP_LABEL


class ExtractionService:
    def __init__(
        self,
        config: dict,
        sevenzip_path: Path | None,
        passwords: list[str],
        messages: Queue[OperationMessage],
    ) -> None:
        self.config = config
        self.sevenzip_path = sevenzip_path
        self.passwords = passwords
        self.messages = messages
        self.stop_event = Event()
        self.thread: Thread | None = None
        self.remembered_password: str | None = None
        self._password_lock = Lock()
        self._progress_lock = Lock()

    @property
    def active(self) -> bool:
        return self.thread is not None and self.thread.is_alive()

    def stop(self) -> None:
        self.stop_event.set()

    def extract_items(self, items: list[QueueItem], test_only: bool = False) -> bool:
        if self.active:
            return False
        self.stop_event.clear()
        self.thread = Thread(target=self._extract_items, args=(items, test_only), name="ExtractorXExtraction", daemon=True)
        self.thread.start()
        return True

    def scan_paths(self, roots: list[Path]) -> bool:
        if self.active:
            return False
        self.stop_event.clear()
        self.thread = Thread(target=self._scan_paths, args=(roots,), name="ExtractorXScan", daemon=True)
        self.thread.start()
        return True

    def _scan_paths(self, roots: list[Path]) -> None:
        _apply_thread_priority(str(self.config.get("ThreadPriority", "Normal")))
        self.messages.put(OperationMessage("scan_started", "Scanning for archives..."))

        def on_path(path: Path) -> None:
            self.messages.put(OperationMessage("archive_found", payload={"path": str(path)}))

        def on_log(text: str, level: str) -> None:
            self.messages.put(OperationMessage("log", text, level=level))

        count = scan_paths(roots, bool(self.config.get("DeepArchiveDetection", True)), self.stop_event, on_path, on_log)
        self.messages.put(OperationMessage("scan_done", f"Scan complete. Found {count} archive(s)."))

    def _extract_items(self, items: list[QueueItem], test_only: bool) -> None:
        try:
            self._extract_items_inner(items, test_only)
        except Exception as exc:  # noqa: BLE001 - top-level guard so UI always sees extract_done
            log.exception("Extraction worker crashed")
            self._log(f"Extraction crashed: {exc}", "error")
            self.messages.put(
                OperationMessage(
                    "extract_done",
                    "Extraction stopped (worker crashed).",
                    level="error",
                    payload={"test_only": test_only},
                )
            )

    def _extract_items_inner(self, items: list[QueueItem], test_only: bool) -> None:
        _apply_thread_priority(str(self.config.get("ThreadPriority", "Normal")))
        priority = str(self.config.get("ThreadPriority", "Normal"))
        if not self.sevenzip_path:
            self.messages.put(OperationMessage("log", "7-Zip was not found. Preparing bundled 7-Zip...", level="warning"))
            try:
                self.sevenzip_path = download_7zip()
                self.messages.put(
                    OperationMessage(
                        "sevenzip_ready",
                        f"7-Zip ready: {self.sevenzip_path}",
                        level="success",
                        payload={"path": str(self.sevenzip_path)},
                    )
                )
            except Exception as exc:
                self.messages.put(OperationMessage("error", f"7-Zip setup failed: {exc}", level="error"))
                self.messages.put(OperationMessage("extract_done", "Extraction stopped."))
                return

        if not check_7zip_version(self.sevenzip_path):
            self._log(
                f"WARNING: 7-Zip is below {MIN_SEVENZIP_LABEL}. "
                f"Update to fix CVE-2025-11001 (symlink RCE) and CVE-2026-48095 (heap overflow).",
                "warning",
            )

        verb = "Testing" if test_only else "Extracting"
        total = len(items)
        self.messages.put(
            OperationMessage(
                "extract_started",
                f"{verb} {total} archive(s)...",
                payload={"total": total, "current": 0, "test_only": test_only},
            )
        )

        allowlist = set(str(ext).lower().lstrip(".") for ext in self.config.get("HandlerAllowlist", []) or [])
        parallelism = max(1, int(self.config.get("MaxParallelExtractions", 1) or 1)) if not test_only else 1
        parallelism = min(parallelism, max(1, len(items)))
        progress_counter = {"value": 0}

        def process(index: int, item: QueueItem) -> None:
            _apply_thread_priority(priority)
            if self.stop_event.is_set():
                return
            with self._progress_lock:
                progress_counter["value"] += 1
                current = progress_counter["value"]
            self.messages.put(
                OperationMessage(
                    "progress",
                    f"{verb} {current}/{total}: {item.archive_path.name}",
                    item_id=item.id,
                    payload={"total": total, "current": current, "test_only": test_only},
                )
            )
            if not item.archive_path.exists():
                item.status = QueueStatus.FAILED
                item.error = "Archive no longer exists."
                self.messages.put(OperationMessage("item_failed", item.error, item_id=item.id, level="error"))
                return
            if is_non_first_volume(item.archive_path):
                item.status = QueueStatus.SKIPPED
                self.messages.put(
                    OperationMessage("item_skipped", "Skipped non-first archive volume.", item_id=item.id, level="warning")
                )
                return
            archive_suffix = item.archive_path.suffix.lstrip(".").lower()
            if allowlist and archive_suffix not in allowlist:
                item.status = QueueStatus.SKIPPED
                self.messages.put(
                    OperationMessage(
                        "item_skipped",
                        f"Skipped: .{archive_suffix} is not in the handler allowlist.",
                        item_id=item.id,
                        level="warning",
                    )
                )
                return

            output = (
                Path(item.output_override).expanduser()
                if item.output_override
                else resolve_output_path(str(self.config.get("OutputPath", "")), item.archive_path)
            )
            item.output_path = output
            if not test_only:
                output.mkdir(parents=True, exist_ok=True)
                run_hook("PreExtractCommand", self.config, item.archive_path, output, self._log)
            self.messages.put(
                OperationMessage(
                    "item_started",
                    f"{verb} {item.archive_path.name}",
                    item_id=item.id,
                    payload={"test_only": test_only},
                )
            )
            retry_max = int(self.config.get("RetryCount", 0) or 0) if not test_only else 0
            retry_delay = int(self.config.get("RetryDelaySeconds", 30) or 30)
            attempt = 0
            while True:
                success, text, password_used = self._try_extract(item.archive_path, output, test_only=test_only)
                if success or attempt >= retry_max or self.stop_event.is_set():
                    break
                attempt += 1
                self._log(f"Retry {attempt}/{retry_max} for {item.archive_path.name} in {retry_delay}s", "warning")
                self.stop_event.wait(retry_delay)
            if success:
                if password_used and bool(self.config.get("AssumeOnePassword", True)):
                    with self._password_lock:
                        self.remembered_password = password_used
                if test_only:
                    item.status = QueueStatus.TEST_OK
                    item.test_detail = "Integrity OK"
                    self.messages.put(
                        OperationMessage(
                            "item_done",
                            f"Test OK: {item.archive_path.name}",
                            item_id=item.id,
                            payload={"test_only": True},
                        )
                    )
                else:
                    escaped = validate_extraction_paths(output)
                    if escaped:
                        for path in escaped[:5]:
                            self._log(f"SECURITY: path escaped output dir: {path}", "error")
                        self._log(
                            f"SECURITY: {len(escaped)} file(s) written outside output directory — possible zip-slip attack.",
                            "error",
                        )
                    output = cleanup_success_output(output, item.archive_path, self.config, self._log)
                    item.output_path = output
                    item.status = QueueStatus.DONE
                    if bool(self.config.get("PropagateMotw", True)):
                        propagate_motw(item.archive_path, output, self._log)
                    self.messages.put(
                        OperationMessage(
                            "item_done",
                            f"Extracted to {output}",
                            item_id=item.id,
                            payload={"output": str(output)},
                        )
                    )
                    if bool(self.config.get("NestedExtraction", True)):
                        self._extract_nested(output, depth=1)
                    run_external_processors(self.config, item.archive_path, output, self._log)
                    run_plugins(item.archive_path, output, self.config, self._log)
                    run_hook("PostExtractCommand", self.config, item.archive_path, output, self._log)
                    apply_post_action(item.archive_path, self.config, self._log)
                    if bool(self.config.get("OpenDestAfterExtract", False)):
                        open_destination(output, self._log)
            else:
                item.status = QueueStatus.FAILED
                item.error = text
                if test_only:
                    item.test_detail = _extract_crc_errors(text) or text
                if not test_only:
                    cleanup_failed_output(item.output_path, self.config, self._log)
                    run_hook("OnFailureCommand", self.config, item.archive_path, output, self._log, exit_code=1)
                self.messages.put(OperationMessage("item_failed", text, item_id=item.id, level="error"))

        if parallelism <= 1:
            for index, item in enumerate(items, start=1):
                if self.stop_event.is_set():
                    break
                try:
                    process(index, item)
                except Exception as exc:  # noqa: BLE001 - isolate per-item failure
                    log.exception("Worker failed for %s", item.archive_path)
                    item.status = QueueStatus.FAILED
                    item.error = f"Worker crashed: {exc}"
                    self.messages.put(
                        OperationMessage("item_failed", item.error, item_id=item.id, level="error")
                    )
        else:
            pool = ThreadPoolExecutor(max_workers=parallelism, thread_name_prefix="ExtractorXWorker")
            try:
                futures = [pool.submit(process, index, item) for index, item in enumerate(items, start=1)]
                # Wait until everything finishes OR a stop is requested. Stop cancels queued
                # futures so they never run, but in-flight workers are allowed to exit cleanly
                # because they check ``stop_event`` frequently.
                while True:
                    done, not_done = wait(futures, timeout=0.25)
                    if self.stop_event.is_set():
                        for future in not_done:
                            future.cancel()
                        break
                    if len(done) == len(futures):
                        break
                for future in futures:
                    if future.done() and not future.cancelled():
                        exc = future.exception()
                        if exc is not None:
                            log.exception("Worker future raised", exc_info=exc)
                            self._log(f"Worker error: {exc}", "error")
            finally:
                pool.shutdown(wait=True, cancel_futures=True)

        done_text = "Test complete." if test_only else "Extraction complete."
        if self.stop_event.is_set():
            done_text = "Stopped."
        self.messages.put(OperationMessage("extract_done", done_text, payload={"test_only": test_only}))

    def _extract_nested(self, folder: Path, depth: int, seen: set[Path] | None = None) -> None:
        max_depth = int(self.config.get("NestedMaxDepth", 5))
        if depth > max_depth or self.stop_event.is_set():
            return
        seen = seen if seen is not None else set()
        candidates: list[Path] = []
        for path in folder.rglob("*"):
            try:
                resolved = path.resolve()
            except OSError:
                continue
            if resolved in seen:
                continue
            if not path.is_file() or is_non_first_volume(path) or not is_supported_archive(path, True):
                continue
            seen.add(resolved)
            candidates.append(path)
        for path in candidates:
            if self.stop_event.is_set():
                break
            if not path.exists():
                continue
            output = path.with_name(archive_name(path))
            output.mkdir(parents=True, exist_ok=True)
            self.messages.put(OperationMessage("log", f"Nested archive: {path.name}"))
            success, text, password_used = self._try_extract(path, output)
            if success:
                if password_used and bool(self.config.get("AssumeOnePassword", True)):
                    with self._password_lock:
                        self.remembered_password = password_used
                cleanup_success_output(output, path, self.config, self._log)
                self.messages.put(OperationMessage("log", f"Nested extracted: {path.name}", level="success"))
                if bool(self.config.get("NestedApplyPostAction", False)):
                    apply_post_action(path, self.config, self._log)
                else:
                    _safe_delete(path)
                self._extract_nested(output, depth + 1, seen)
            else:
                cleanup_failed_output(output, self.config, self._log)
                self.messages.put(OperationMessage("log", f"Nested failed: {path.name}: {text}", level="error"))

    def _check_decompression_ratio(self, archive: Path, output_text: str) -> str | None:
        """Return an error message if the decompression ratio exceeds the configured limit."""
        max_ratio = int(self.config.get("MaxDecompressionRatio", 1000) or 0)
        if max_ratio <= 0:
            return None
        try:
            archive_size = archive.stat().st_size
        except OSError:
            return None
        if archive_size <= 0:
            return None
        extracted_size = 0
        for line in output_text.splitlines():
            match = _re.search(r"Size:\s+(\d+)", line)
            if match:
                extracted_size = max(extracted_size, int(match.group(1)))
        if extracted_size <= 0:
            return None
        ratio = extracted_size / archive_size
        if ratio > max_ratio:
            return (
                f"Decompression ratio {ratio:.0f}:1 exceeds limit {max_ratio}:1 "
                f"({archive.name}: {archive_size} bytes -> {extracted_size} bytes). "
                f"Possible zip bomb."
            )
        return None

    def _try_extract(self, archive: Path, output: Path, test_only: bool = False) -> tuple[bool, str, str | None]:
        with self._password_lock:
            remembered = self.remembered_password
        skip_after = int(self.config.get("SkipAfterFailedPasswords", 0) or 0)
        sidecar = load_sidecar_passwords(archive) if bool(self.config.get("UsePasswordSidecars", True)) else []
        if sidecar:
            self._log(f"Loaded {len(sidecar)} sidecar password(s) for {archive.name}", "info")
        attempts = build_password_attempts(
            remembered_password=remembered,
            saved_passwords=self.passwords,
            use_password_list=bool(self.config.get("UsePasswordList", True)),
            skip_after_failures=skip_after,
            sidecar_passwords=sidecar,
            wordlist=bool(self.config.get("WordlistGeneration", False)),
            wordlist_max=int(self.config.get("WordlistMaxAttempts", 500) or 500),
            password_rules=list(self.config.get("PasswordRules", []) or []),
            archive_name=archive.name,
        )
        use_hash_probe = bool(self.config.get("HashModePasswordProbe", True))
        encoding_setting = str(self.config.get("FilenameEncoding", "Auto"))
        detected_encoding: str | None = None
        if encoding_setting == "Auto":
            detected_encoding = detect_zip_codepage(archive)
            if detected_encoding:
                self._log(f"Auto-detected codepage {detected_encoding} for {archive.name}", "info")

        # When there are multiple password candidates and we are extracting (not
        # just testing), use a fast ``7z t`` probe to find the correct password
        # first, then do a single verbose extraction with the winner. This is
        # 5-10x faster than full-extraction probing on large/solid archives.
        password_candidates = [p for p in attempts if p is not None]
        if use_hash_probe and not test_only and len(password_candidates) > 1:
            winning_password = self._probe_password(archive, attempts, detected_encoding=detected_encoding)
            if winning_password is not None:
                # Replace the attempt list with just the no-password probe and
                # the winning password for the real extraction pass.
                attempts = [winning_password]
            elif winning_password is None and self.stop_event.is_set():
                return False, "Cancelled.", None

        last_error = ""
        for password in attempts:
            if self.stop_event.is_set():
                return False, "Cancelled.", None
            success, text, code, cancelled = self._run_sevenzip(
                archive, output, password, test_only=test_only, verbose=True,
                detected_encoding=detected_encoding,
            )
            if cancelled:
                return False, "Cancelled.", None
            if code in {0, 1}:
                bomb_msg = self._check_decompression_ratio(archive, text)
                if bomb_msg and not test_only:
                    self._log(f"SECURITY: {bomb_msg}", "error")
                return True, text, password
            last_error = text or f"7-Zip exited with code {code}."
        if self.passwords or sidecar:
            return False, last_error or "Extraction failed after trying saved passwords.", None
        return False, last_error or "Extraction failed.", None

    def _probe_password(
        self, archive: Path, attempts: list[str | None],
        detected_encoding: str | None = None,
    ) -> str | None:
        """Use fast ``7z t`` to find the first working password.

        Returns the winning password string, or ``None`` if no password
        matched (the caller should fall back to the full attempt list).
        The initial ``None`` (no-password) probe is included so unencrypted
        archives short-circuit immediately.
        """
        for password in attempts:
            if self.stop_event.is_set():
                return None
            success, text, code, cancelled = self._run_sevenzip(
                archive, archive.parent, password, test_only=True, verbose=False,
                detected_encoding=detected_encoding,
            )
            if cancelled:
                return None
            if code in {0, 1}:
                if password is not None:
                    self._log(f"Password probe succeeded for {archive.name}", "info")
                return password
        return None

    def _run_sevenzip(
        self,
        archive: Path,
        output: Path,
        password: str | None,
        *,
        test_only: bool = False,
        verbose: bool = True,
        detected_encoding: str | None = None,
    ) -> tuple[bool, str, int, bool]:
        """Execute a single 7-Zip invocation.

        Returns ``(success, tail_text, exit_code, cancelled)``.
        """
        effective_encoding = str(self.config.get("FilenameEncoding", "Auto"))
        if effective_encoding == "Auto" and detected_encoding:
            effective_encoding = detected_encoding
        command = build_sevenzip_command(
            sevenzip_path=self.sevenzip_path,
            archive=archive,
            output=output,
            overwrite_mode=str(self.config.get("OverwriteMode", "Always")),
            exclusions=str(self.config.get("FileExclusions", "")),
            inclusions=str(self.config.get("IncludeMasks", "")),
            password=password,
            test_only=test_only,
            filename_encoding=effective_encoding,
        )
        try:
            proc = subprocess.Popen(
                command,
                stdout=subprocess.PIPE,
                stderr=subprocess.STDOUT,
                text=True,
                encoding="utf-8",
                errors="replace",
                creationflags=subprocess.CREATE_NO_WINDOW if hasattr(subprocess, "CREATE_NO_WINDOW") else 0,
            )
        except OSError as exc:
            return False, str(exc), -1, False
        output_lines: list[str] = []
        tail_limit = 200
        cancelled = False
        try:
            assert proc.stdout is not None
            for line in proc.stdout:
                if self.stop_event.is_set():
                    cancelled = True
                    break
                clean = line.strip()
                if clean:
                    output_lines.append(clean)
                    if len(output_lines) > tail_limit:
                        del output_lines[:-tail_limit]
                    if verbose:
                        self.messages.put(OperationMessage("log", clean))
                    pct_match = _re.match(r"(\d+)%", clean)
                    if pct_match and verbose:
                        self.messages.put(
                            OperationMessage("sub_progress", payload={"percent": int(pct_match.group(1))})
                        )
        finally:
            if cancelled:
                _terminate_process(proc)
            try:
                code = proc.wait(timeout=30)
            except subprocess.TimeoutExpired:
                _terminate_process(proc, kill=True)
                try:
                    code = proc.wait(timeout=5)
                except subprocess.TimeoutExpired:
                    code = -1
            if proc.stdout is not None:
                try:
                    proc.stdout.close()
                except OSError:
                    pass
        text = "\n".join(output_lines[-8:])
        return code in {0, 1}, text, code, cancelled

    def _log(self, text: str, level: str = "info") -> None:
        self.messages.put(OperationMessage("log", text, level=level))


def load_sidecar_passwords(archive: Path) -> list[str]:
    """Read passwords from sidecar files next to *archive*.

    Checks for ``<archive>.pwd.txt`` and ``passwords.txt`` in the archive's
    parent directory. Each non-empty line is treated as a candidate password.
    Missing or unreadable files are silently skipped.
    """
    passwords: list[str] = []
    candidates = [
        archive.with_name(archive.name + ".pwd.txt"),
        archive.parent / "passwords.txt",
    ]
    for path in candidates:
        try:
            if not path.is_file():
                continue
            text = path.read_text(encoding="utf-8-sig", errors="replace")
            for line in text.splitlines():
                stripped = line.strip()
                if stripped and stripped not in passwords:
                    passwords.append(stripped)
        except OSError:
            continue
    return passwords


def build_password_attempts(
    remembered_password: str | None,
    saved_passwords: list[str],
    use_password_list: bool = True,
    skip_after_failures: int = 0,
    sidecar_passwords: list[str] | None = None,
    wordlist: bool = False,
    wordlist_max: int = 500,
    password_rules: list[dict] | None = None,
    archive_name: str = "",
) -> list[str | None]:
    """Return the ordered password attempts used by the extraction loop.

    Mirrors the legacy PowerShell flow: try without a password first so that
    unencrypted archives complete without pointless retries, then fall through
    to sidecar passwords (from ``<archive>.pwd.txt`` / ``passwords.txt``),
    the remembered batch password (if one has been learned), and the saved
    password list.

    When *wordlist* is ``True``, the stored password list is expanded with
    case, leet-speak, and date-suffix permutations (capped at *wordlist_max*).

    ``skip_after_failures`` caps the number of *password* attempts (not
    counting the initial no-password probe). ``0`` disables the cap.
    """
    attempts: list[str | None] = [None]
    # Sidecar passwords are the highest-priority candidates after the
    # no-password probe, since they are co-located with the archive and
    # likely to be the correct password.
    for password in sidecar_passwords or []:
        if password and password not in attempts:
            attempts.append(password)
    for rule in password_rules or []:
        pattern = str(rule.get("Pattern", ""))
        rule_passwords = rule.get("Passwords", [])
        if pattern and archive_name:
            try:
                if _re.search(pattern, archive_name, _re.IGNORECASE):
                    for pw in rule_passwords:
                        if pw and pw not in attempts:
                            attempts.append(pw)
            except _re.error:
                continue
    if remembered_password and remembered_password not in attempts:
        attempts.append(remembered_password)
    effective_passwords = list(saved_passwords)
    if wordlist and effective_passwords:
        effective_passwords = generate_wordlist(effective_passwords, max_total=wordlist_max)
    if use_password_list:
        for password in effective_passwords:
            if password and password not in attempts:
                attempts.append(password)
    if skip_after_failures > 0:
        # Preserve the initial ``None`` probe and cap the trailing password list.
        password_attempts = [attempt for attempt in attempts if attempt is not None]
        if len(password_attempts) > skip_after_failures:
            password_attempts = password_attempts[:skip_after_failures]
        attempts = [None, *password_attempts]
    return attempts


def _terminate_process(proc: subprocess.Popen[str], *, kill: bool = False) -> None:
    """Best-effort termination of a Popen child.

    Uses :meth:`subprocess.Popen.terminate` first (``SIGTERM`` / ``TerminateProcess``)
    then upgrades to :meth:`kill` if ``kill`` is requested. Never raises; used from
    cleanup paths where a secondary failure would mask the original problem.
    """
    try:
        if kill:
            proc.kill()
        else:
            proc.terminate()
    except (OSError, ValueError):
        pass


def _safe_delete(path: Path) -> None:
    try:
        if path.is_file():
            path.unlink()
        elif path.is_dir():
            shutil.rmtree(path)
    except OSError:
        pass


_CRC_PATTERNS = _re.compile(
    r"(?i)(CRC\s*Error|Data\s*Error|checksum\s*error|Headers\s*Error|"
    r"Unexpected\s*end\s*of\s*archive|Can\s*not\s*open\s*.*?as\s*archive|"
    r"ERROR:\s*Data\s*Error|Sub\s*items\s*Errors)",
)


def _extract_crc_errors(output_text: str) -> str:
    """Parse 7-Zip test output and return a summary of CRC/integrity errors.

    If the output contains recognizable integrity-failure markers, they are
    gathered into a short summary string. Returns an empty string when no
    patterns match.
    """
    lines = output_text.strip().splitlines()
    errors: list[str] = []
    for line in lines:
        if _CRC_PATTERNS.search(line):
            errors.append(line.strip())
    return "; ".join(errors[:10]) if errors else ""


def _apply_thread_priority(priority: str) -> None:
    if os.name != "nt":
        return
    values = {
        "Low": -2,
        "BelowNormal": -1,
        "Normal": 0,
        "AboveNormal": 1,
        "High": 2,
    }
    try:
        kernel32 = ctypes.windll.kernel32
        handle = kernel32.GetCurrentThread()
        kernel32.SetThreadPriority(handle, values.get(priority, 0))
    except Exception:
        pass

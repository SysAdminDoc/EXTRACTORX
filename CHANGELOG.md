# Changelog

All notable changes to EXTRACTORX will be documented in this file.

## [v2.3.1] - 2026-04-22 -- Hardening audit

Comprehensive production-hardening pass across the whole codebase. No new
end-user features; focus is on correctness, data safety, and resilience.

### Fixed -- Data safety
- **Atomic config writes.** `save_config` now writes to a sibling temp file,
  fsyncs, and `os.replace`'s into place. An interrupted write can no longer
  corrupt `config.json`.
- **Atomic password store writes.** `PasswordStore.save` uses the same
  temp+rename pattern. DPAPI-protected blobs can no longer be half-written.
- **Corrupt config recovery.** `load_config` now backs up malformed or
  non-object JSON to `config.corrupt-<hex>.json` and logs a warning before
  falling back to defaults -- users can recover their previous content.
- **Diagnostic logging on decryption failure.** `PasswordStore.load` no longer
  swallows DPAPI/JSON errors silently; the cause is written to the logger.
- **Atomic 7-Zip bootstrap download.** `download_7zip` streams to a `.part`
  file and atomically replaces the target, so an aborted download cannot leave
  a truncated `7zr.exe` behind.

### Fixed -- Concurrency and reliability
- **Parallel cancel no longer deadlocks.** The `ThreadPoolExecutor` loop
  cancels queued futures the moment a stop is requested and always emits
  exactly one `extract_done` -- the UI never gets stuck in `Stopping...`.
- **Top-level worker guard.** `_extract_items` wraps the real body in a
  try/except that guarantees `extract_done` fires even if a worker crashes.
- **Process leak on cancel closed.** `_try_extract` now uses a
  try/finally that terminates the 7-Zip child, waits with a timeout,
  escalates to kill, and closes `proc.stdout` on every path.
- **`_extract_nested` no longer races on `remembered_password`**; the
  password lock is taken consistently with the outer loop.
- **Worker threads honor `ThreadPriority`.** `_apply_thread_priority` is now
  called per worker, not just on the dispatcher thread.
- **Parallelism clamped to `len(items)`** to avoid spinning up unused workers.

### Fixed -- Input validation and safety
- **Download helper** now validates scheme+host, caps transfers at 4 GiB,
  honors a 120 s default timeout, sends a User-Agent header, sanitizes the
  target filename, checks `Content-Length` up-front, and aborts the stream
  if the running byte count exceeds the cap.
- **Watch-folder queue** validates each detected file via
  `is_supported_archive` before enqueueing, so non-archive files dropped into
  a watched folder can no longer pollute the queue.
- **Scanner + watcher** were rewritten around `os.scandir` with
  per-directory error isolation (permission errors on one subtree no longer
  crash the scan), and they skip directory symlinks / Windows junctions to
  prevent infinite cycles.
- **`recycle_path` double-null termination** now uses an explicit
  `create_unicode_buffer` and an `absolute-path` resolve; avoids the subtle
  `c_wchar_p` truncation footgun.
- **Scan-a-drive-root** now triggers a confirmation dialog so users don't
  accidentally kick off a `C:\` scan.

### Fixed -- Shell and file-association hygiene
- **File associations back up the previous default ProgID** on install and
  restore it on uninstall, so removing ExtractorX associations returns the
  user's original handler (e.g. 7-Zip) instead of orphaning the extension.

### Fixed -- UI resilience
- **Crash-safe message pump.** `_pump_messages` stops scheduling itself once
  the window is closing and swallows `TclError`s during teardown so
  late-arriving messages never raise.
- **Close path is idempotent** and tolerates destroyed widgets (geometry,
  output-template read, watcher stop, shell-bridge dispose each wrapped).
- **Retry Selected** only targets items in retriable states; running items
  are skipped with a log line.
- **Password retry prompt** filters out items that have since completed and
  uses the entropy-meter dialog instead of the plain `simpledialog`.
- **Export Script** now uses proper cmd.exe quoting via the shared
  `_cmd_quote` helper and emits CRLF line endings for Windows compatibility.
- **Sevenzip-ready / archive-found messages** tolerate empty payloads and
  skip duplicates instead of blindly queueing.

### Changed
- `_portable_root()` result is cached after the first call -- every
  `app_data_dir()` lookup no longer re-iterates candidate paths.
- `download_7zip` uses an explicit `subprocess.TimeoutExpired` handler and
  no longer imports `subprocess` inside a `try/except`.
- Dead `_ = sys` / `_ = os` sentinel imports removed from `app.py` and
  `download.py`.

### Tests
Added 8 new cases covering atomic-write behavior, corrupt-config recovery,
password-store round-trip, download-scheme rejection, scanner isolation,
drive-root detection, and a full end-to-end parallel-cancel regression
test. **58 tests passing**, including the existing parity suite.

## [v2.3.0] - 2026-04-22

Inspired by a competitive scan across ExtractNow, PeaZip, NanaZip, 7-Zip ZS,
Bandizip, WinRAR, Ark, dtrx, patool, The Unarchiver, UniExtract 2, QArchive,
Unpackerr, and Microsoft's RecursiveExtractor.

### Added
- **Smart Extract policy** (`Auto` / `AlwaysWrap` / `NeverWrap`) -- the legacy
  duplicate-folder cleanup is now one of three explicit modes. `NeverWrap`
  unwraps any single-child wrapper (tarbomb-safe), `AlwaysWrap` preserves
  7-Zip's layout verbatim.
- **Filename encoding override** -- Settings > Destination lets you pin 7-Zip
  to a specific code page (UTF-8, cp437, cp932, cp936, cp949, cp950, cp1251,
  cp1252) for archives produced in other locales.
- **Include masks** -- companion to File Exclusions; only files matching the
  mask list are extracted (emits `-ir!` switches).
- **Parallel extraction** -- Advanced > "Parallel extractions" (1-8) runs
  concurrent 7-Zip processes with a shared lock on the remembered-password
  slot so cached credentials still propagate safely.
- **Lifecycle hooks** -- pre-extract, post-extract, and on-failure shell
  commands with `{ArchivePath}`, `{Output}`, `{ArchiveName}`, `{ExitCode}`
  tokens (auto-quoted).
- **Backend selector** -- `SevenZipOverride` config key + Advanced field lets
  you point ExtractorX at 7-Zip ZS (`7zzs.exe`), NanaZip (`NanaZipC.exe`), or
  any other 7-Zip compatible binary.
- **Handler allowlist** -- Settings > Files accepts a whitelist of archive
  extensions; everything else is skipped with a "not in handler allowlist"
  status.
- **Skip-after-N failed passwords** policy prevents runaway cycling through
  large saved-password lists.
- **Password entropy meter** -- the Add Password dialog shows a
  weak/fair/strong label plus Shannon-style bit estimate.
- **Bookmarks** -- Settings > Bookmarks and a toolbar dropdown. Bookmarked
  folders launch a scan; bookmarked archives queue directly.
- **Identify button + `--identify` CLI** -- reports each file's format and
  support status without extracting.
- **Export Script** toolbar button emits a `.cmd` that replays the current
  queue.
- **URL argument support** -- startup paths beginning with `http://` /
  `https://` download into `%APPDATA%/ExtractorX/downloads` before queueing.
- **Portable mode** -- if a sibling `portable.flag` exists next to the
  entrypoint, config/logs/passwords live in `ExtractorX.data/` next to the
  app instead of `%APPDATA%`. The window title notes "(portable)".

### Changed
- `find_7zip()` now also searches for `7zzs.exe` and NanaZip's
  `NanaZipC.exe` on PATH and accepts an explicit override argument.
- `build_sevenzip_command()` gained `inclusions=` and `filename_encoding=`
  parameters; existing callers continue to work unchanged.

### Fixed
- ResourceWarning from `_try_extract` -- `proc.stdout` is now closed
  explicitly after the read loop.

## [v2.2.3] - 2026-04-22

### Added
- Header shows a live "Watching N folder(s)" indicator whenever the watch
  service is running.
- Queue right-click menu gained `Retry Selected` (requeue failed/finished
  items and re-run extraction) and `Set Destination...` (pick a per-item
  override folder with one file dialog for the whole selection).

### Changed
- Eligible-for-extraction status set now includes `PasswordRequired` so items
  parked for a password prompt can be re-run once credentials exist.

## [v2.2.2] - 2026-04-22

### Added
- Window title now reports the running ExtractorX version.
- Queue right-click menu gained a "Copy Destination Path" entry.
- Settings dialog has a "Reset to Defaults" button that preserves window
  geometry, watch folders, and the password list.
- External-processor token expansion: archive/destination tokens are now
  auto-quoted for `cmd.exe`, any user-placed wrapping quotes are absorbed so
  `"{ArchivePath}"` and `{ArchivePath}` behave identically, and embedded
  quotes are escaped as `""`.

### Changed
- Log pane is capped at the last 2000 lines to keep long-running sessions
  from growing the Tk Text buffer unbounded; the on-disk history log is
  unaffected.
- `extractorx/__init__.py` `__version__` set to `"2.2.1"` so UI/title/docs
  stay in sync. (Bumped further with this release.)

## [v2.2.1] - 2026-04-22

### Added
- Double-click on a queued row now opens the destination (if extracted) or
  reveals the source archive.
- Log pane right-click menu (Copy, Clear Log, Open Log Folder).
- File-association delta handling: extensions removed from the Settings list
  are now uninstalled from the registry on save.
- Off-screen window guard: saved window positions outside the current
  screen geometry are ignored so the window always lands on-screen.
- Automatic migration of any legacy `passwords.py.dat` file to the correct
  `passwords.dat` filename on first launch.

### Changed
- Nested extraction iterates a snapshot of archive paths and tracks them in a
  resolved `seen` set so extracted children cannot be re-discovered mid-walk.
- Stop button now reports "Nothing to stop." when idle and sets the header
  status to "Stopping..." while the active job winds down; status resets to
  "7-Zip ready" once `extract_done` arrives.
- Progress bar resets to zero at the start of each batch.
- 7-Zip output buffer is capped at the last 200 lines to keep long-running
  extractions bounded.

## [v2.2.0] - 2026-04-22

### Added
- Modular Python port under `extractorx/` with a Tkinter UI, polling-based watch
  service, DPAPI-backed password store, native Windows drag/drop and tray bridge,
  Explorer context-menu integration, and a sortable queue.
- `{Program Files}` / `{ProgramFiles}` output-path macro for parity with the
  legacy PowerShell app's token set.
- Password list import from a plain text file in Settings > Passwords.
- Live theme reload from the Settings dialog (no more restart prompt).
- `tests/` unit suite covering config normalization, output macros, 7-Zip
  command layout, password-attempt ordering, post-processing cleanup, shell
  integration helpers, and a live watch-folder round trip.

### Changed
- Password attempt order now tries "no password" first, then the remembered
  batch password, then the saved list -- matching the legacy PowerShell flow.
- Watched archives are now batched into a single extraction call per pump,
  avoiding churn when bursts of files land in a watched folder.
- `ExtractorX.Legacy.ps1` preserves the polished PowerShell script without
  modification.

### Fixed
- Password store filename corrected from `passwords.py.dat` to `passwords.dat`.
- `cleanup_failed_output` no longer reports a bogus deletion when the path was
  already absent.
- `apply_post_action` with an empty `PostActionFolder` now refuses the move
  instead of silently using the current working directory.
- Drag/drop filter no longer applies to files added via the `Add Files` dialog.
- Tkinter popup menus release their grab in a `finally` block, preventing
  stuck modal input.
- Watcher dedup now resolves paths and prunes stale `_seen` entries so the
  in-memory map cannot grow unbounded.

## [v2.1.0]

- Original polished PowerShell release (preserved as `ExtractorX.Legacy.ps1`).

# ROADMAP

Backlog for EXTRACTORX. The WPF app and Python port both ship today; this captures the feature
gap versus 7-Zip, NanaZip, PeaZip, Bandizip, and what would push EXTRACTORX past ExtractNow.

## Planned Features

### Core extraction
- **Test-archive batch mode** — verify integrity across the queue without extracting; surface CRC
  errors in a dedicated results tab.
- **Solid-archive streaming extract** — for `.7z` / `.rar`, stream specific files out without
  decompressing the whole solid block when the backend supports it.
- **Compressed-archive creation parity** — match the "Compress" side of PeaZip/Bandizip so one tool
  handles both directions; start with ZIP/7z/TAR.zst, reuse the password DPAPI store.
- **Format auto-registration toggle per extension** — granular control in Settings > Explorer so a
  user can own `.7z` but leave `.zip` with Explorer's built-in handler.
- **Pattern-scoped extraction** — `--include-glob` / `--exclude-glob` at the CLI and a filter box
  in the extract dialog, so large archives can be partially extracted without manual browsing.

### Password workflow
- **Hash-mode password probe** — use `7z t -p<candidate>` with early abort to cycle candidates 5-10x
  faster than the current full-extract probe on large archives.
- **Wordlist generation** — combine the stored password list with case/leet permutations and
  date-suffix variants, cap attempts, log every try for audit.
- **Password sidecar files** — auto-read `<archive>.pwd.txt` or `passwords.txt` from the archive's
  directory before falling back to the global list.

### Automation
- **Watch-folder rules per path** — different output templates, post-actions, and password lists
  per watched directory.
- **Quarantine on suspicious archive** — optional pre-extract scan via MSMpEng `MpCmdRun.exe` or
  ClamAV, move flagged archives to a quarantine folder.
- **Scheduled extraction** — Task Scheduler integration so large/slow jobs can run overnight.

### UI
- **Archive browser tab** — open an archive as a virtual tree, drag individual files out without
  extracting the whole thing (PeaZip-style).
- **Hex view** for small files pulled out of archives (integrity spot-check without a separate
  viewer).
- **Dockable log pane** with structured columns (timestamp, archive, file, level) and CSV export.
- **High-contrast theme** variant for accessibility compliance.

### Platform
- **Winget manifest** + **Scoop bucket entry** for one-line install.
- **Code-signing** the PowerShell host exe and Python exe; embed timestamped Authenticode so
  SmartScreen stops flagging fresh builds.

## Competitive Research

- **7-Zip / NanaZip** — the reference implementations. Matching feature parity isn't the goal;
  batch UX and watch-folder automation remain EXTRACTORX's edge.
- **PeaZip** — 200+ formats, archive browser as virtual tree, built-in converter. The virtual-tree
  browser is the single biggest missing feature here.
- **Bandizip** — fastest extraction benchmarks, "Extract Here (Smart)" auto-wrap. EXTRACTORX
  already has Smart Extract policy; validate perf against Bandizip on large `.7z` / `.rar`.
- **ExtractNow (original)** — no longer actively developed; replace its workflow feature-for-feature
  then keep going.

## Nice-to-Haves

- **Cloud source adapters** — extract directly from S3/GDrive/OneDrive URLs without manual download.
- **Mount-as-drive** via Dokan for browsing large archives without extracting (WinFsp fallback).
- **Archive diff** — compare two versions of the same archive, show added/removed/changed files.
- **WSL backend option** — shell out to `7z` / `unar` inside WSL for exotic Unix formats (`.xar`,
  `.sit`, `.z`) that Windows ports handle poorly.
- **Portable flag propagation to PowerShell app** — the Python port already supports
  `portable.flag`; wire the WPF app to the same convention.
- **Plugin SDK** for custom post-processors (Python entry points, PowerShell modules).

## Open-Source Research (Round 2)

### Related OSS Projects
- **7-Zip** — https://github.com/ip7z/7zip — the reference archive engine; LZMA/LZMA2, 7z format; what EXTRACTORX already calls into
- **PeaZip** — https://github.com/peazip/PeaZip — Free-Pascal GUI wrapping 7-Zip + several engines; broad format matrix (ARJ, ACE, CAB, ISO, WIM) close to EXTRACTORX's list
- **The Unarchiver (macOS)** — https://github.com/MacPaw/The-Unarchiver — reference for password-prompt UX and legacy-format support (StuffIt, ARJ, LZH)
- **p7zip** — https://github.com/p7zip-project/p7zip — Unix port; community test corpus useful for regression
- **Archive Extractor (Python)** — https://github.com/gdraheim/zzipdoc — not primary but demonstrates Python-only extraction patterns
- **libarchive** — https://github.com/libarchive/libarchive — C library behind bsdtar; EXTRACTORX's Python port could call `python-libarchive-c` for extended format coverage
- **Hydrus Network** — https://github.com/hydrusnetwork/hydrus — watches folders, imports, dedupes — reference for the "watch & extract" daemon loop
- **Hazel (closed) / Watson-style folder-rules OSS: directory-monitor** — https://github.com/daandelange/directory-watcher patterns
- **patool** — https://github.com/wummel/patool — Python multi-backend CLI archive wrapper; good unified-interface pattern for the Python port
- **bandizip** / **WinRAR** — closed but reference for ergonomic drag-drop and silent-install defaults

### Features to Borrow
- libarchive-backed Python port for 15+ more obscure formats (WARC, XAR, LHA, ALZ, MTZ) without shelling out to 7-Zip (python-libarchive-c)
- Format plugin registry where each archive type declares extract/list/test hooks, and a common adapter wires them in (patool pattern) — swap 7z/libarchive/rar at runtime
- "Smart defaults" recipe per-format: ISO → mount vs extract toggle, deb/rpm → extract metadata + data.tar.* cleanly, msi → extract files + streams (PeaZip behavior)
- One-shot "repair archive" pass using PAR2 when a `.par2` sidecar is present (PeaZip + multipar integration)
- Batch hash-compare: verify every extracted file against embedded SFV/CRC32/SHA1 lists in RAR/7z (7-Zip `t` command) — surface a failed-integrity list
- CLI mode for each GUI action with a logged exit-equivalent command shown in the log panel (UX power-user pattern)
- Output folder templates with tokens: `{parent}\{archive_name_no_ext}\`, `{date}\{archive_name}\`, "flat if single-folder inside" (PeaZip)
- Watch-folder daemon (optional service) that extracts newly-dropped archives into a mirrored target tree, configurable password list scope per watch (Hydrus-style import folder)
- Shell-integration menu with per-item actions: Extract Here / Extract To Subfolder / Extract and Test / Extract and Delete original (all 3 competitors have this)
- Signed-release + SHA256SUMS + Sigstore/cosign attestation for binaries — matches ExtractNow's gap (general best practice)
- Drop targets on the frameless window: separate target for "extract here" vs "extract to subfolder" (The Unarchiver and PeaZip both do this)

### Patterns & Architectures Worth Studying
- Engine abstraction: GUI calls a stable "archive engine" interface; 7-Zip, libarchive, native RAR lib plug in behind it (patool, PeaZip)
- Format detection by magic bytes with a deterministic fallback order — never trust extensions (libarchive does this exclusively; EXTRACTORX already has "deep detection", validate parity)
- Multi-volume detection state machine: recognize `.rar/.r00/.r01`, `.7z.001`, `.part01.rar`, `.z01/.zip` splits and only kick off once (7-Zip behavior, PeaZip polish)
- Password cycling with per-archive persistent salt: tried-passwords hash set per archive so a retry doesn't re-try known-bad ones (PeaZip)
- DPAPI-encrypted local password store with per-user rotation on Windows — this is a strong EXTRACTORX pattern; ensure the Python port uses `win32crypt.CryptProtectData` equivalent via `pywin32` to match parity

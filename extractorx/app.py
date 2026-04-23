"""Application bootstrap and dependency wiring."""

from __future__ import annotations

import argparse
from pathlib import Path

from .archive import is_archive_path, is_supported_archive
from .config import is_portable, load_config
from .download import download_archive, looks_like_url
from .identify import identify
from .models import QueueItem
from .passwords import PasswordStore
from .sevenzip import find_7zip


def build_parser() -> argparse.ArgumentParser:
    parser = argparse.ArgumentParser(description="ExtractorX archive extraction app")
    parser.add_argument("files", nargs="*", help="Archives, folders, or URLs to queue at startup")
    parser.add_argument("--target", dest="target_path", help="Destination override")
    parser.add_argument("--extract-here", action="store_true", help="Extract each archive to its containing folder")
    parser.add_argument("--auto-extract", action="store_true", help="Start extracting queued startup archives immediately")
    parser.add_argument("--test", action="store_true", help="Test queued startup archives instead of extracting")
    parser.add_argument("--scan", action="store_true", help="Treat startup directories as folders to scan for archives")
    parser.add_argument("--minimize", action="store_true", help="Start minimized")
    parser.add_argument("--minimizetotray", action="store_true", help="Start minimized")
    parser.add_argument(
        "--identify",
        action="store_true",
        help="Print archive format identification for the given paths and exit without launching the UI",
    )
    return parser


def _run_identify(paths: list[str]) -> int:
    if not paths:
        print("--identify requires at least one path")
        return 2
    for raw in paths:
        result = identify(Path(raw).expanduser())
        status = "supported" if result.supported else "unsupported"
        print(f"{result.path}\t{result.format}\t{status}\t{result.reason}")
    return 0


def main(argv: list[str] | None = None) -> int:
    args = build_parser().parse_args(argv)
    if args.identify:
        return _run_identify(args.files)

    config = load_config()
    sevenzip_path = find_7zip(override=config.get("SevenZipOverride"))
    password_store = PasswordStore()

    startup_items: list[QueueItem] = []
    startup_scan_paths: list[Path] = []
    for raw_path in args.files:
        raw_path = str(raw_path).strip()
        if not raw_path:
            continue
        if looks_like_url(raw_path):
            downloaded = download_archive(raw_path, log=lambda text, level="info": print(f"[{level}] {text}"))
            if downloaded and is_supported_archive(
                downloaded, deep_detection=bool(config.get("DeepArchiveDetection", True))
            ):
                output_override = str(downloaded.parent) if args.extract_here else args.target_path
                startup_items.append(QueueItem.from_path(downloaded, output_override=output_override))
            elif downloaded:
                print(f"[warning] Downloaded file is not a supported archive: {downloaded}")
            continue
        path = Path(raw_path).expanduser()
        if path.exists() and path.is_file() and is_archive_path(path):
            output_override = str(path.parent) if args.extract_here else args.target_path
            startup_items.append(QueueItem.from_path(path, output_override=output_override))
        elif path.exists() and path.is_dir():
            startup_scan_paths.append(path)

    from .ui import ExtractorXApp  # imported lazily so --identify never initializes Tk

    app = ExtractorXApp(
        config=config,
        sevenzip_path=sevenzip_path,
        password_store=password_store,
        startup_items=startup_items,
        startup_scan_paths=startup_scan_paths,
        auto_extract_startup=bool(args.auto_extract),
        test_startup=bool(args.test),
        start_minimized=args.minimize or args.minimizetotray,
        portable=is_portable(),
    )
    app.run()
    return 0

"""Core smoke tests for the Python port."""

from __future__ import annotations

import os
import tempfile
import threading
import unittest
from pathlib import Path
from queue import Queue

from extractorx.archive import archive_name, detect_zip_codepage, has_archive_magic, is_non_first_volume, resolve_output_path, validate_extraction_paths
from extractorx.config import DEFAULT_CONFIG, normalize_config
from extractorx.extractor import build_password_attempts
from extractorx.models import QueueItem, QueueStatus
from extractorx.postprocess import (
    apply_post_action,
    cleanup_failed_output,
    cleanup_success_output,
    expand_processor_command,
)
from extractorx.sevenzip import build_sevenzip_command, overwrite_switch
from extractorx.shell_integration import (
    _association_prog_id,
    _command,
    _normalize_extensions,
)
from extractorx.watcher import WatchService


class ConfigNormalizationTests(unittest.TestCase):
    def test_clamps_and_defaults(self) -> None:
        config = normalize_config(
            {
                "Theme": "White",
                "NestedMaxDepth": "999",
                "PasswordTimeout": "bad",
                "WatchFolders": "C:/Temp",
            }
        )
        self.assertEqual(config["Theme"], "White")
        self.assertEqual(config["NestedMaxDepth"], 50)
        self.assertEqual(config["PasswordTimeout"], 45)
        self.assertEqual(config["WatchFolders"], ["C:/Temp"])

    def test_rejects_invalid_enum_values(self) -> None:
        config = normalize_config(
            {
                "Theme": "Clownshoes",
                "OverwriteMode": "whatever",
                "PostAction": "Shred",
                "DragDropFilterType": "Nope",
                "ThreadPriority": "Ludicrous",
            }
        )
        self.assertEqual(config["Theme"], "Midnight")
        self.assertEqual(config["OverwriteMode"], "Always")
        self.assertEqual(config["PostAction"], "None")
        self.assertEqual(config["DragDropFilterType"], "None")
        self.assertEqual(config["ThreadPriority"], "Normal")

    def test_normalizes_external_processors(self) -> None:
        raw = [
            {"Extension": ".RAR", "Command": "notepad {ArchivePath}"},
            {"Extension": "", "Command": "no-ext"},
            {"Extension": "zip", "Command": ""},
            "garbage",
            {"Extension": "7z", "Command": "test"},
        ]
        config = normalize_config({"ExternalProcessors": raw})
        self.assertEqual(
            config["ExternalProcessors"],
            [
                {"Extension": "RAR", "Command": "notepad {ArchivePath}"},
                {"Extension": "7z", "Command": "test"},
            ],
        )

    def test_bool_coercion(self) -> None:
        config = normalize_config({"MinimizeToTray": "no", "SoundsEnabled": "1", "AlwaysOnTop": 0})
        self.assertFalse(config["MinimizeToTray"])
        self.assertTrue(config["SoundsEnabled"])
        self.assertFalse(config["AlwaysOnTop"])

    def test_defaults_match_declared_keys(self) -> None:
        config = normalize_config(None)
        for key in DEFAULT_CONFIG:
            self.assertIn(key, config, msg=f"missing default for {key}")


class ArchiveHelperTests(unittest.TestCase):
    def test_archive_name_handles_compound_suffixes(self) -> None:
        self.assertEqual(archive_name(Path("backup.tar.gz")), "backup")
        self.assertEqual(archive_name(Path("BACKUP.TAR.GZ")), "BACKUP")
        self.assertEqual(archive_name(Path("archive.tgz")), "archive")
        self.assertEqual(archive_name(Path("plain.zip")), "plain")

    def test_output_path_uses_literal_replacements(self) -> None:
        resolved = resolve_output_path(r"{ArchiveFolder}\{ArchiveName}", Path(r"C:\Temp\demo.zip"))
        self.assertTrue(str(resolved).endswith(r"C:\Temp\demo"))

    def test_output_path_expands_env_and_program_files(self) -> None:
        os.environ["EXTRACTORX_TEST_TOKEN"] = "hello"
        try:
            resolved = resolve_output_path(
                r"{Env:EXTRACTORX_TEST_TOKEN}\{ArchiveName}",
                Path(r"C:\Temp\demo.zip"),
            )
            self.assertTrue(str(resolved).endswith(r"hello\demo"))
            resolved_pf = resolve_output_path(r"{Program Files}\demo", Path(r"C:\Temp\demo.zip"))
            self.assertIn("Program Files", str(resolved_pf))
        finally:
            os.environ.pop("EXTRACTORX_TEST_TOKEN", None)

    def test_unique_archive_name_macro(self) -> None:
        with tempfile.TemporaryDirectory() as folder:
            root = Path(folder)
            archive = root / "demo.zip"
            archive.write_text("zip", encoding="utf-8")
            (root / "demo").mkdir()
            resolved = resolve_output_path(str(root / "{ArchiveNameUnique}"), archive)
            self.assertEqual(resolved.name, "demo (1)")

    def test_unique_archive_collision_walks_counters(self) -> None:
        with tempfile.TemporaryDirectory() as folder:
            root = Path(folder)
            archive = root / "demo.zip"
            archive.write_text("zip", encoding="utf-8")
            for suffix in ("demo", "demo (1)", "demo (2)"):
                (root / suffix).mkdir()
            resolved = resolve_output_path(str(root / "{ArchiveNameUnique}"), archive)
            self.assertEqual(resolved.name, "demo (3)")

    def test_multi_volume_detection(self) -> None:
        self.assertTrue(is_non_first_volume(Path("movie.part02.rar")))
        self.assertFalse(is_non_first_volume(Path("movie.part01.rar")))
        self.assertTrue(is_non_first_volume(Path("pack.r01")))
        self.assertTrue(is_non_first_volume(Path("split.bin.002")))
        self.assertFalse(is_non_first_volume(Path("split.bin.001")))

    def test_zip_magic_detection(self) -> None:
        with tempfile.TemporaryDirectory() as folder:
            path = Path(folder) / "archive.bin"
            path.write_bytes(bytes.fromhex("504B0304") + b"body")
            self.assertTrue(has_archive_magic(path))

    def test_missing_file_does_not_crash_magic_detection(self) -> None:
        self.assertFalse(has_archive_magic(Path("does/not/exist.bin")))


class SevenZipCommandTests(unittest.TestCase):
    def test_extract_command_layout(self) -> None:
        command = build_sevenzip_command(
            sevenzip_path=Path("C:/Program Files/7-Zip/7z.exe"),
            archive=Path("C:/tmp/demo.zip"),
            output=Path("C:/out/demo"),
            overwrite_mode="Always",
            exclusions="Thumbs.db;desktop.ini",
            password=None,
            test_only=False,
        )
        self.assertEqual(command[1], "x")
        self.assertEqual(command[2], str(Path("C:/tmp/demo.zip")))
        self.assertEqual(command[3], f"-o{Path('C:/out/demo')}")
        self.assertIn("-y", command)
        self.assertIn("-aoa", command)
        self.assertIn("-bb1", command)
        self.assertIn("-xr!Thumbs.db", command)
        self.assertIn("-xr!desktop.ini", command)
        self.assertFalse(any(flag.startswith("-p") for flag in command))

    def test_test_mode_omits_output_switch(self) -> None:
        command = build_sevenzip_command(
            sevenzip_path=Path("7z"),
            archive=Path("demo.zip"),
            output=Path("C:/out/demo"),
            test_only=True,
        )
        self.assertEqual(command[1], "t")
        self.assertFalse(any(flag.startswith("-o") for flag in command))

    def test_password_is_appended_when_provided(self) -> None:
        command = build_sevenzip_command(
            sevenzip_path=Path("7z"),
            archive=Path("demo.zip"),
            output=Path("demo"),
            password="s3cr3t",
        )
        self.assertIn("-ps3cr3t", command)

    def test_overwrite_switch_mapping(self) -> None:
        self.assertEqual(overwrite_switch("Always"), "-aoa")
        self.assertEqual(overwrite_switch("Never"), "-aos")
        self.assertEqual(overwrite_switch("Rename"), "-aou")
        self.assertEqual(overwrite_switch("unknown"), "-aoa")


class PasswordAttemptTests(unittest.TestCase):
    def test_order_prefers_no_password_first(self) -> None:
        self.assertEqual(build_password_attempts(None, []), [None])
        self.assertEqual(build_password_attempts(None, ["alpha", "beta"]), [None, "alpha", "beta"])

    def test_remembered_password_follows_no_password(self) -> None:
        self.assertEqual(
            build_password_attempts("remembered", ["alpha"]),
            [None, "remembered", "alpha"],
        )

    def test_saved_list_respected_for_dedup(self) -> None:
        self.assertEqual(
            build_password_attempts("alpha", ["alpha", "beta"]),
            [None, "alpha", "beta"],
        )

    def test_password_list_can_be_disabled(self) -> None:
        self.assertEqual(
            build_password_attempts("alpha", ["beta"], use_password_list=False),
            [None, "alpha"],
        )


class PostProcessTests(unittest.TestCase):
    def test_cleanup_and_move_action(self) -> None:
        messages: list[tuple[str, str]] = []

        def log(text: str, level: str = "info") -> None:
            messages.append((text, level))

        with tempfile.TemporaryDirectory() as folder:
            root = Path(folder)
            archive = root / "sample.zip"
            archive.write_text("zip", encoding="utf-8")
            output = root / "sample"
            inner = output / "sample"
            inner.mkdir(parents=True)
            (inner / "file.txt").write_text("body", encoding="utf-8")
            cleanup_success_output(
                output,
                archive,
                {"RemoveDuplicateFolder": True, "RenameSingleFile": True},
                log,
            )
            self.assertTrue((output / "sample.txt").exists())

            broken = root / "broken"
            broken.mkdir()
            (broken / "partial.txt").write_text("bad", encoding="utf-8")
            cleanup_failed_output(broken, {"DeleteBrokenFiles": True}, log)
            self.assertFalse(broken.exists())

            apply_post_action(
                archive,
                {"PostAction": "MoveToFolder", "PostActionFolder": str(root / "done")},
                log,
            )
            self.assertTrue((root / "done" / "sample.zip").exists())

    def test_move_action_refuses_empty_destination(self) -> None:
        messages: list[tuple[str, str]] = []

        def log(text: str, level: str = "info") -> None:
            messages.append((text, level))

        with tempfile.TemporaryDirectory() as folder:
            archive = Path(folder) / "sample.zip"
            archive.write_text("zip", encoding="utf-8")
            apply_post_action(archive, {"PostAction": "MoveToFolder", "PostActionFolder": ""}, log)
            self.assertTrue(archive.exists(), "archive must stay put when folder is blank")
            self.assertTrue(any(level == "warning" for _, level in messages))

    def test_cleanup_failed_output_ignores_missing(self) -> None:
        logs: list[tuple[str, str]] = []
        cleanup_failed_output(
            Path(tempfile.gettempdir()) / "extractorx-does-not-exist",
            {"DeleteBrokenFiles": True},
            lambda text, level="info": logs.append((text, level)),
        )
        self.assertEqual(logs, [])

    def test_expand_processor_command_returns_list(self) -> None:
        template = "notepad {ArchivePath}"
        expanded = expand_processor_command(
            template,
            archive=Path(r"C:\My Stuff\demo.zip"),
            output=Path(r"D:\out\demo"),
        )
        self.assertIsInstance(expanded, list)
        self.assertEqual(expanded[0], "notepad")
        self.assertTrue(len(expanded) >= 2)

    def test_expand_processor_command_absorbs_wrapping_quotes(self) -> None:
        template = 'notepad "{ArchivePath}"'
        expanded = expand_processor_command(
            template,
            archive=Path(r"C:\path\file.zip"),
            output=Path(r"C:\out"),
        )
        self.assertIsInstance(expanded, list)
        self.assertEqual(expanded[0], "notepad")

    def test_expand_processor_command_handles_spaces(self) -> None:
        expanded = expand_processor_command(
            "tool {ArchivePath}",
            archive=Path(r"C:\My Files\test.zip"),
            output=Path(r"C:\out"),
        )
        self.assertIsInstance(expanded, list)
        self.assertTrue(any("My Files" in arg for arg in expanded))

    def test_duplicate_folder_flatten_handles_trailing_noise(self) -> None:
        with tempfile.TemporaryDirectory() as folder:
            root = Path(folder)
            archive = root / "sample.zip"
            archive.write_text("zip", encoding="utf-8")
            output = root / "sample"
            inner = output / "sample"
            inner.mkdir(parents=True)
            (inner / "a.txt").write_text("a", encoding="utf-8")
            (inner / "b.txt").write_text("b", encoding="utf-8")
            logs: list[tuple[str, str]] = []
            cleanup_success_output(
                output,
                archive,
                {"RemoveDuplicateFolder": True, "RenameSingleFile": False},
                lambda text, level="info": logs.append((text, level)),
            )
            self.assertTrue((output / "a.txt").exists())
            self.assertTrue((output / "b.txt").exists())


class PasswordStoreTests(unittest.TestCase):
    def test_round_trip(self) -> None:
        from extractorx.passwords import PasswordStore

        with tempfile.TemporaryDirectory() as folder:
            store = PasswordStore(path=Path(folder) / "passwords.dat")
            store.save(["alpha", "beta", "alpha", "", "gamma"])
            loaded = store.load()
            self.assertEqual(loaded, ["alpha", "beta", "gamma"])

    def test_missing_file_returns_empty_list(self) -> None:
        from extractorx.passwords import PasswordStore

        with tempfile.TemporaryDirectory() as folder:
            store = PasswordStore(path=Path(folder) / "missing.dat")
            self.assertEqual(store.load(), [])

    def test_migrates_legacy_filename(self) -> None:
        from extractorx.passwords import PasswordStore

        with tempfile.TemporaryDirectory() as folder:
            target = Path(folder) / "passwords.dat"
            legacy = Path(folder) / "passwords.py.dat"
            seed = PasswordStore(path=legacy)
            seed.save(["migrated"])
            self.assertTrue(legacy.exists())
            store = PasswordStore(path=target)
            self.assertTrue(target.exists())
            self.assertFalse(legacy.exists())
            self.assertEqual(store.load(), ["migrated"])


class ExtractionServiceTests(unittest.TestCase):
    def test_parallel_cancellation_emits_extract_done(self) -> None:
        """With parallelism > 1, a stop signal still produces exactly one extract_done.

        Regression for the earlier hang where ThreadPoolExecutor's ``__exit__``
        waited for in-flight futures but the enclosing loop broke out before
        sending ``extract_done``, leaving the UI in a ``Stopping...`` state.
        """
        import time
        from queue import Queue

        from extractorx.extractor import ExtractionService
        from extractorx.models import OperationMessage, QueueItem
        from extractorx.sevenzip import find_7zip

        sevenzip = find_7zip()
        if not sevenzip:
            self.skipTest("7-Zip is not available in this environment")
        with tempfile.TemporaryDirectory() as folder:
            root = Path(folder)
            archives = []
            import zipfile

            for i in range(3):
                path = root / f"p{i}.zip"
                with zipfile.ZipFile(path, "w") as zf:
                    zf.writestr(f"x{i}.txt", "x")
                archives.append(path)
            output_dir = root / "out"
            cfg = {
                "OutputPath": str(output_dir / "{ArchiveName}"),
                "OverwriteMode": "Always",
                "FileExclusions": "",
                "IncludeMasks": "",
                "NestedExtraction": False,
                "UsePasswordList": False,
                "AssumeOnePassword": True,
                "RemoveDuplicateFolder": True,
                "RenameSingleFile": False,
                "ThreadPriority": "Normal",
                "DeleteAfterExtract": False,
                "PostAction": "None",
                "OpenDestAfterExtract": False,
                "DeleteBrokenFiles": False,
                "MaxParallelExtractions": 2,
                "SmartExtract": "Auto",
                "FilenameEncoding": "Auto",
                "SkipAfterFailedPasswords": 0,
                "HandlerAllowlist": [],
                "PreExtractCommand": "",
                "PostExtractCommand": "",
                "OnFailureCommand": "",
            }
            items = [QueueItem.from_path(p) for p in archives]
            messages: Queue[OperationMessage] = Queue()
            service = ExtractionService(cfg, sevenzip, [], messages)
            service.extract_items(items, test_only=False)
            service.stop()
            deadline = time.time() + 10
            while service.active and time.time() < deadline:
                time.sleep(0.05)
            self.assertFalse(service.active, "Service never finished after stop")
            # Drain messages and confirm exactly one extract_done arrived.
            done_count = 0
            while not messages.empty():
                msg = messages.get()
                if msg.type == "extract_done":
                    done_count += 1
            self.assertEqual(done_count, 1)

    def test_output_override_takes_precedence_over_template(self) -> None:
        import zipfile
        from queue import Queue

        from extractorx.extractor import ExtractionService
        from extractorx.models import QueueItem, QueueStatus
        from extractorx.sevenzip import find_7zip

        sevenzip = find_7zip()
        if not sevenzip:
            self.skipTest("7-Zip is not available in this environment")
        with tempfile.TemporaryDirectory() as folder:
            root = Path(folder)
            archive = root / "payload.zip"
            with zipfile.ZipFile(archive, "w") as zf:
                zf.writestr("payload.txt", "content")
            override = root / "custom"
            config = {
                "OutputPath": str(root / "ignored" / "{ArchiveName}"),
                "OverwriteMode": "Always",
                "FileExclusions": "",
                "NestedExtraction": False,
                "UsePasswordList": False,
                "AssumeOnePassword": True,
                "RemoveDuplicateFolder": True,
                "RenameSingleFile": False,
                "ThreadPriority": "Normal",
                "DeleteAfterExtract": False,
                "PostAction": "None",
                "OpenDestAfterExtract": False,
                "DeleteBrokenFiles": False,
            }
            item = QueueItem.from_path(archive, output_override=str(override))
            messages: Queue = Queue()
            service = ExtractionService(config, sevenzip, [], messages)
            service.extract_items([item], test_only=False)
            import time

            deadline = time.time() + 15
            while service.active and time.time() < deadline:
                time.sleep(0.1)
            self.assertFalse(service.active, "extraction never finished")
            self.assertEqual(item.status, QueueStatus.DONE)
            self.assertTrue((override / "payload.txt").exists(), "override destination was not honored")
            self.assertFalse((root / "ignored" / "payload" / "payload.txt").exists())


class SmartExtractTests(unittest.TestCase):
    def test_auto_mode_removes_duplicate_archive_name(self) -> None:
        with tempfile.TemporaryDirectory() as folder:
            root = Path(folder)
            archive = root / "sample.zip"
            archive.write_text("zip", encoding="utf-8")
            output = root / "sample"
            (output / "sample" / "a.txt").parent.mkdir(parents=True)
            (output / "sample" / "a.txt").write_text("x", encoding="utf-8")
            logs: list[tuple[str, str]] = []
            cleanup_success_output(
                output,
                archive,
                {"SmartExtract": "Auto", "RemoveDuplicateFolder": True},
                lambda text, level="info": logs.append((text, level)),
            )
            self.assertTrue((output / "a.txt").exists())

    def test_never_wrap_flattens_any_single_child(self) -> None:
        from extractorx.postprocess import cleanup_success_output as cleanup

        with tempfile.TemporaryDirectory() as folder:
            root = Path(folder)
            archive = root / "misc.zip"
            archive.write_text("zip", encoding="utf-8")
            output = root / "misc"
            inner = output / "unrelated-wrapper" / "payload.txt"
            inner.parent.mkdir(parents=True)
            inner.write_text("payload", encoding="utf-8")
            logs: list[tuple[str, str]] = []
            cleanup(
                output,
                archive,
                {"SmartExtract": "NeverWrap"},
                lambda text, level="info": logs.append((text, level)),
            )
            self.assertTrue((output / "payload.txt").exists())

    def test_always_wrap_preserves_layout(self) -> None:
        from extractorx.postprocess import cleanup_success_output as cleanup

        with tempfile.TemporaryDirectory() as folder:
            root = Path(folder)
            archive = root / "keep.zip"
            archive.write_text("zip", encoding="utf-8")
            output = root / "keep"
            inner = output / "keep" / "file.txt"
            inner.parent.mkdir(parents=True)
            inner.write_text("value", encoding="utf-8")
            cleanup(output, archive, {"SmartExtract": "AlwaysWrap"}, lambda *_: None)
            self.assertTrue((output / "keep" / "file.txt").exists())


class SevenZipEncodingTests(unittest.TestCase):
    def test_filename_encoding_emits_code_page(self) -> None:
        command = build_sevenzip_command(
            sevenzip_path="7z",
            archive="a.zip",
            output="out",
            filename_encoding="cp932",
        )
        self.assertIn("-mcp=932", command)

    def test_auto_encoding_omits_code_page(self) -> None:
        command = build_sevenzip_command(
            sevenzip_path="7z",
            archive="a.zip",
            output="out",
            filename_encoding="Auto",
        )
        self.assertFalse(any(flag.startswith("-mcp=") for flag in command))

    def test_include_masks_emit_ir_flags(self) -> None:
        command = build_sevenzip_command(
            sevenzip_path="7z",
            archive="a.zip",
            output="out",
            inclusions="*.json;*.md",
        )
        self.assertIn("-ir!*.json", command)
        self.assertIn("-ir!*.md", command)


class PasswordAttemptPolicyTests(unittest.TestCase):
    def test_skip_after_caps_password_count(self) -> None:
        attempts = build_password_attempts(
            remembered_password="remembered",
            saved_passwords=["alpha", "beta", "gamma"],
            skip_after_failures=2,
        )
        self.assertEqual(attempts, [None, "remembered", "alpha"])

    def test_skip_after_zero_is_unlimited(self) -> None:
        attempts = build_password_attempts(
            remembered_password=None,
            saved_passwords=["a", "b", "c", "d"],
            skip_after_failures=0,
        )
        self.assertEqual(attempts, [None, "a", "b", "c", "d"])


class PasswordEntropyTests(unittest.TestCase):
    def test_weak_fair_strong(self) -> None:
        from extractorx.passwords import classify_entropy, estimate_entropy_bits

        self.assertEqual(classify_entropy("")[0], "weak")
        self.assertEqual(classify_entropy("abc")[0], "weak")
        self.assertLess(estimate_entropy_bits("abc"), estimate_entropy_bits("Abc123"))
        self.assertEqual(classify_entropy("CorrectHorseBatteryStaple!9")[0], "strong")


class HandlerAllowlistTests(unittest.TestCase):
    def test_normalizes_allowlist_entries(self) -> None:
        from extractorx.config import normalize_config

        config = normalize_config({"HandlerAllowlist": [" .ZIP ", "7z", "", "RAR"]})
        self.assertEqual(config["HandlerAllowlist"], ["zip", "7z", "rar"])


class BookmarkConfigTests(unittest.TestCase):
    def test_normalizes_string_and_dict_bookmarks(self) -> None:
        from extractorx.config import normalize_config

        config = normalize_config(
            {
                "Bookmarks": [
                    {"Label": " Pictures ", "Path": " C:/Users/me/Pictures "},
                    "literal",
                    {"Label": "", "Path": "ignored"},
                ]
            }
        )
        self.assertEqual(
            config["Bookmarks"],
            [
                {"Label": "Pictures", "Path": "C:/Users/me/Pictures"},
                {"Label": "literal", "Path": "literal"},
            ],
        )


class IdentifyTests(unittest.TestCase):
    def test_zip_magic_is_identified(self) -> None:
        from extractorx.identify import identify

        with tempfile.TemporaryDirectory() as folder:
            path = Path(folder) / "payload.bin"
            path.write_bytes(bytes.fromhex("504B0304") + b"more")
            result = identify(path)
            self.assertTrue(result.supported)
            self.assertEqual(result.format, "ZIP")

    def test_unknown_reports_not_supported(self) -> None:
        from extractorx.identify import identify

        with tempfile.TemporaryDirectory() as folder:
            path = Path(folder) / "plain.txt"
            path.write_text("hello", encoding="utf-8")
            result = identify(path)
            self.assertFalse(result.supported)
            self.assertEqual(result.format, "unknown")


class DownloadHelperTests(unittest.TestCase):
    def test_looks_like_url(self) -> None:
        from extractorx.download import looks_like_url

        self.assertTrue(looks_like_url("https://example.com/a.zip"))
        self.assertTrue(looks_like_url("http://example.com/a.zip"))
        self.assertFalse(looks_like_url("C:/Users/x.zip"))
        self.assertFalse(looks_like_url("ftp://example.com/a.zip"))
        self.assertFalse(looks_like_url(""))
        self.assertFalse(looks_like_url("not a url at all"))

    def test_download_rejects_unsupported_scheme(self) -> None:
        from extractorx.download import download_archive

        logs: list[tuple[str, str]] = []
        result = download_archive("ftp://example.com/a.zip", log=lambda t, l="info": logs.append((t, l)))
        self.assertIsNone(result)
        self.assertTrue(any("Refusing" in message for message, _ in logs))

    def test_download_empty_url_rejected(self) -> None:
        from extractorx.download import download_archive

        self.assertIsNone(download_archive("", log=None))


class AtomicWriteTests(unittest.TestCase):
    def test_config_write_is_atomic(self) -> None:
        from extractorx.config import _atomic_write

        with tempfile.TemporaryDirectory() as folder:
            target = Path(folder) / "config.json"
            _atomic_write(target, '{"hello": "world"}')
            self.assertEqual(target.read_text(encoding="utf-8"), '{"hello": "world"}')
            # No stray temp files left behind.
            stragglers = [p for p in Path(folder).iterdir() if p != target]
            self.assertEqual(stragglers, [])

    def test_config_load_backs_up_corrupt_file(self) -> None:
        import extractorx.config as config_module

        with tempfile.TemporaryDirectory() as folder:
            root = Path(folder)
            target = root / "config.json"
            target.write_text("{ not valid json", encoding="utf-8")
            original = config_module.config_path
            config_module.config_path = lambda: target  # type: ignore[assignment]
            try:
                config = config_module.load_config()
            finally:
                config_module.config_path = original  # type: ignore[assignment]
            self.assertIn("Theme", config)
            corrupt = list(root.glob("config.corrupt-*.json"))
            self.assertGreaterEqual(len(corrupt), 1)


class PasswordAtomicSaveTests(unittest.TestCase):
    def test_save_is_atomic_and_idempotent(self) -> None:
        from extractorx.passwords import PasswordStore

        with tempfile.TemporaryDirectory() as folder:
            store = PasswordStore(path=Path(folder) / "passwords.dat")
            store.save(["alpha", "beta"])
            self.assertEqual(store.load(), ["alpha", "beta"])
            # Writing twice should leave no stray temp files around.
            store.save(["alpha", "beta", "gamma"])
            stragglers = [p for p in Path(folder).iterdir() if not p.name.endswith("passwords.dat")]
            self.assertEqual(stragglers, [])
            self.assertEqual(store.load(), ["alpha", "beta", "gamma"])


class ScannerRobustnessTests(unittest.TestCase):
    def test_scanner_reports_matches_and_skips_nonexistent(self) -> None:
        import zipfile
        from threading import Event

        from extractorx.scanner import scan_paths

        with tempfile.TemporaryDirectory() as folder:
            root = Path(folder)
            nested = root / "sub" / "deep"
            nested.mkdir(parents=True)
            archive = nested / "found.zip"
            with zipfile.ZipFile(archive, "w") as zf:
                zf.writestr("x.txt", "x")
            discovered: list[Path] = []
            count = scan_paths(
                [root, root / "does_not_exist"],
                deep_detection=True,
                should_stop=Event(),
                on_path=discovered.append,
            )
            self.assertEqual(count, 1)
            self.assertEqual(discovered, [archive])


class DrivelookTests(unittest.TestCase):
    def test_is_drive_root(self) -> None:
        from extractorx.ui import _is_drive_root

        self.assertTrue(_is_drive_root(Path("C:/")))
        self.assertFalse(_is_drive_root(Path(tempfile.gettempdir())))


class HookDispatchTests(unittest.TestCase):
    def test_empty_hook_is_noop(self) -> None:
        from extractorx.hooks import run_hook

        logs: list[tuple[str, str]] = []
        run_hook(
            "PreExtractCommand",
            {"PreExtractCommand": ""},
            Path("irrelevant.zip"),
            Path("."),
            lambda text, level="info": logs.append((text, level)),
        )
        self.assertEqual(logs, [])


class ShellIntegrationTests(unittest.TestCase):
    def test_command_helpers(self) -> None:
        self.assertEqual(_association_prog_id(".zip"), "ExtractorX.zip")
        self.assertEqual(_normalize_extensions(["zip", ".7Z", "zip"]), [".zip", ".7z"])
        command = _command(Path("C:/repo/ExtractorX.py"), "--test")
        self.assertIn("--test", command)
        self.assertIn('"%1"', command)

    def test_command_substitutes_background_target_token(self) -> None:
        command = _command(Path("C:/repo/ExtractorX.py"), "--scan", "%V")
        self.assertIn('"%V"', command)
        self.assertNotIn('"%1"', command)
        self.assertIn("--scan", command)


class WatcherTests(unittest.TestCase):
    def test_new_archive_triggers_queue(self) -> None:
        with tempfile.TemporaryDirectory() as folder:
            watched = Path(folder)
            queue: Queue[Path] = Queue()
            service = WatchService([str(watched)], queue, deep_detection=True)
            service.start()
            try:
                archive = watched / "drop.zip"
                archive.write_bytes(bytes.fromhex("504B0304") + b"payload")
                deadline = threading.Event()
                deadline.wait(0.2)
                found_path: Path | None = None
                for _ in range(60):
                    try:
                        found_path = queue.get(timeout=0.5)
                        break
                    except Exception:
                        if not service.active:
                            break
                self.assertIsNotNone(found_path, "watcher did not observe the new archive")
                self.assertEqual(found_path.name, "drop.zip")
            finally:
                service.stop()


class SidecarPasswordTests(unittest.TestCase):
    def test_loads_archive_pwd_file(self) -> None:
        from extractorx.extractor import load_sidecar_passwords

        with tempfile.TemporaryDirectory() as folder:
            root = Path(folder)
            archive = root / "secret.7z"
            archive.write_text("placeholder", encoding="utf-8")
            sidecar = root / "secret.7z.pwd.txt"
            sidecar.write_text("pass1\npass2\n", encoding="utf-8")
            passwords = load_sidecar_passwords(archive)
            self.assertEqual(passwords, ["pass1", "pass2"])

    def test_loads_directory_passwords_file(self) -> None:
        from extractorx.extractor import load_sidecar_passwords

        with tempfile.TemporaryDirectory() as folder:
            root = Path(folder)
            archive = root / "archive.zip"
            archive.write_text("placeholder", encoding="utf-8")
            (root / "passwords.txt").write_text("abc\ndef\n", encoding="utf-8")
            passwords = load_sidecar_passwords(archive)
            self.assertEqual(passwords, ["abc", "def"])

    def test_deduplicates_and_skips_empty_lines(self) -> None:
        from extractorx.extractor import load_sidecar_passwords

        with tempfile.TemporaryDirectory() as folder:
            root = Path(folder)
            archive = root / "dup.zip"
            archive.write_text("placeholder", encoding="utf-8")
            sidecar = root / "dup.zip.pwd.txt"
            sidecar.write_text("alpha\n\nalpha\nbeta\n", encoding="utf-8")
            passwords = load_sidecar_passwords(archive)
            self.assertEqual(passwords, ["alpha", "beta"])

    def test_missing_sidecars_returns_empty(self) -> None:
        from extractorx.extractor import load_sidecar_passwords

        with tempfile.TemporaryDirectory() as folder:
            archive = Path(folder) / "nopasswords.zip"
            archive.write_text("placeholder", encoding="utf-8")
            self.assertEqual(load_sidecar_passwords(archive), [])

    def test_build_password_attempts_with_sidecars(self) -> None:
        attempts = build_password_attempts(
            remembered_password=None,
            saved_passwords=["saved1"],
            sidecar_passwords=["sidecar1", "sidecar2"],
        )
        # Sidecar passwords should come before saved passwords
        self.assertEqual(attempts, [None, "sidecar1", "sidecar2", "saved1"])

    def test_sidecar_dedup_with_remembered(self) -> None:
        attempts = build_password_attempts(
            remembered_password="sidecar1",
            saved_passwords=["saved1"],
            sidecar_passwords=["sidecar1"],
        )
        self.assertEqual(attempts, [None, "sidecar1", "saved1"])


class WordlistTests(unittest.TestCase):
    def test_generates_case_variants(self) -> None:
        from extractorx.passwords import generate_wordlist

        result = generate_wordlist(["Hello"], case_variants=True, leet_variants=False, date_suffixes=False)
        self.assertIn("Hello", result)
        self.assertIn("hello", result)
        self.assertIn("HELLO", result)
        self.assertIn("hELLO", result)

    def test_generates_leet_variants(self) -> None:
        from extractorx.passwords import generate_wordlist

        result = generate_wordlist(["test"], case_variants=False, leet_variants=True, date_suffixes=False)
        self.assertIn("test", result)
        self.assertIn("7357", result)

    def test_generates_date_suffixes(self) -> None:
        from datetime import datetime
        from extractorx.passwords import generate_wordlist

        result = generate_wordlist(["pw"], case_variants=False, leet_variants=False, date_suffixes=True)
        year = str(datetime.now().year)
        self.assertIn("pw" + year, result)
        self.assertIn("pw!", result)
        self.assertIn("pw123", result)

    def test_respects_max_total(self) -> None:
        from extractorx.passwords import generate_wordlist

        result = generate_wordlist(["a", "b", "c"], max_total=5)
        self.assertLessEqual(len(result), 5)
        self.assertEqual(result[:3], ["a", "b", "c"])

    def test_empty_input(self) -> None:
        from extractorx.passwords import generate_wordlist

        self.assertEqual(generate_wordlist([]), [])

    def test_build_password_attempts_with_wordlist(self) -> None:
        attempts = build_password_attempts(
            remembered_password=None,
            saved_passwords=["pass"],
            wordlist=True,
            wordlist_max=50,
        )
        self.assertIn(None, attempts)
        self.assertIn("pass", attempts)
        # Should have variants beyond just the original
        self.assertGreater(len(attempts), 2)


class HashModeProbeTests(unittest.TestCase):
    def test_run_sevenzip_method_exists(self) -> None:
        """Verify the refactored _run_sevenzip helper is callable."""
        from extractorx.extractor import ExtractionService

        self.assertTrue(hasattr(ExtractionService, "_run_sevenzip"))
        self.assertTrue(hasattr(ExtractionService, "_probe_password"))


class HighContrastThemeTests(unittest.TestCase):
    def test_high_contrast_theme_exists(self) -> None:
        from extractorx.themes import get_theme

        palette = get_theme("HighContrast")
        self.assertEqual(palette["bg"], "#000000")
        self.assertEqual(palette["text"], "#FFFFFF")
        self.assertEqual(palette["accent"], "#FFFF00")

    def test_high_contrast_in_config_themes(self) -> None:
        from extractorx.config import THEMES

        self.assertIn("HighContrast", THEMES)

    def test_config_accepts_high_contrast(self) -> None:
        config = normalize_config({"Theme": "HighContrast"})
        self.assertEqual(config["Theme"], "HighContrast")


class PatternScopedExtractionTests(unittest.TestCase):
    def test_include_glob_cli_argument(self) -> None:
        from extractorx.app import build_parser

        parser = build_parser()
        args = parser.parse_args(["--include-glob", "*.json;*.md", "archive.zip"])
        self.assertEqual(args.include_glob, "*.json;*.md")

    def test_exclude_glob_cli_argument(self) -> None:
        from extractorx.app import build_parser

        parser = build_parser()
        args = parser.parse_args(["--exclude-glob", "Thumbs.db;*.tmp", "archive.zip"])
        self.assertEqual(args.exclude_glob, "Thumbs.db;*.tmp")


class NewConfigKeysTests(unittest.TestCase):
    def test_new_defaults_present(self) -> None:
        config = normalize_config(None)
        self.assertTrue(config["UsePasswordSidecars"])
        self.assertTrue(config["HashModePasswordProbe"])
        self.assertFalse(config["WordlistGeneration"])
        self.assertEqual(config["WordlistMaxAttempts"], 500)

    def test_wordlist_max_clamped(self) -> None:
        config = normalize_config({"WordlistMaxAttempts": 99999})
        self.assertEqual(config["WordlistMaxAttempts"], 10000)
        config = normalize_config({"WordlistMaxAttempts": 1})
        self.assertEqual(config["WordlistMaxAttempts"], 10)


class ZipSlipValidationTests(unittest.TestCase):
    def test_no_escapes_returns_empty(self) -> None:
        with tempfile.TemporaryDirectory() as tmp:
            output = Path(tmp) / "output"
            output.mkdir()
            (output / "file.txt").write_text("ok")
            (output / "sub").mkdir()
            (output / "sub" / "nested.txt").write_text("ok")
            self.assertEqual(validate_extraction_paths(output), [])

    def test_detects_symlink_escape(self) -> None:
        if os.name != "nt":
            self.skipTest("symlink test is Windows-specific")
        with tempfile.TemporaryDirectory() as tmp:
            output = Path(tmp) / "output"
            output.mkdir()
            (output / "safe.txt").write_text("ok")
            escaped = validate_extraction_paths(output)
            self.assertEqual(escaped, [])


class DecompressionRatioTests(unittest.TestCase):
    def test_default_present_in_config(self) -> None:
        config = normalize_config(None)
        self.assertEqual(config["MaxDecompressionRatio"], 1000)

    def test_ratio_clamped(self) -> None:
        config = normalize_config({"MaxDecompressionRatio": 999999})
        self.assertEqual(config["MaxDecompressionRatio"], 100000)
        config = normalize_config({"MaxDecompressionRatio": -5})
        self.assertEqual(config["MaxDecompressionRatio"], 0)


class SevenZipVersionTests(unittest.TestCase):
    def test_min_version_constants(self) -> None:
        from extractorx.sevenzip import MIN_SEVENZIP_VERSION, MIN_SEVENZIP_LABEL
        self.assertEqual(MIN_SEVENZIP_VERSION, (26, 2))
        self.assertEqual(MIN_SEVENZIP_LABEL, "26.02")


class SystemThemeTests(unittest.TestCase):
    def test_detect_returns_dark_or_light(self) -> None:
        from extractorx.windows_integration import detect_system_theme
        result = detect_system_theme()
        self.assertIn(result, ("dark", "light"))


class MotwConfigTests(unittest.TestCase):
    def test_propagate_motw_default_true(self) -> None:
        config = normalize_config(None)
        self.assertTrue(config["PropagateMotw"])

    def test_propagate_motw_normalizes_bool(self) -> None:
        config = normalize_config({"PropagateMotw": "false"})
        self.assertFalse(config["PropagateMotw"])


class CodepageDetectionTests(unittest.TestCase):
    def test_returns_none_for_nonzip(self) -> None:
        with tempfile.TemporaryDirectory() as tmp:
            path = Path(tmp) / "test.7z"
            path.write_bytes(b"\x37\x7a\xbc\xaf\x27\x1c")
            self.assertIsNone(detect_zip_codepage(path))

    def test_returns_none_for_utf8_zip(self) -> None:
        import zipfile
        with tempfile.TemporaryDirectory() as tmp:
            path = Path(tmp) / "test.zip"
            with zipfile.ZipFile(path, "w") as zf:
                zf.writestr("readme.txt", "hello")
                zf.writestr("data/output.csv", "a,b,c")
            self.assertIsNone(detect_zip_codepage(path))

    def test_returns_none_for_missing_file(self) -> None:
        self.assertIsNone(detect_zip_codepage(Path("/nonexistent/test.zip")))


class TaskbarProgressTests(unittest.TestCase):
    def test_taskbar_constants_defined(self) -> None:
        from extractorx.windows_integration import TBPF_NOPROGRESS, TBPF_NORMAL, TBPF_ERROR
        self.assertEqual(TBPF_NOPROGRESS, 0)
        self.assertEqual(TBPF_NORMAL, 2)
        self.assertEqual(TBPF_ERROR, 4)


class RetryConfigTests(unittest.TestCase):
    def test_retry_defaults(self) -> None:
        config = normalize_config(None)
        self.assertEqual(config["RetryCount"], 0)
        self.assertEqual(config["RetryDelaySeconds"], 30)

    def test_retry_clamped(self) -> None:
        config = normalize_config({"RetryCount": 99, "RetryDelaySeconds": 9999})
        self.assertEqual(config["RetryCount"], 10)
        self.assertEqual(config["RetryDelaySeconds"], 600)


class SecureDeleteConfigTests(unittest.TestCase):
    def test_default_false(self) -> None:
        config = normalize_config(None)
        self.assertFalse(config["SecureDelete"])


class PasswordRulesTests(unittest.TestCase):
    def test_regex_match_prioritizes_rule_passwords(self) -> None:
        attempts = build_password_attempts(
            remembered_password=None,
            saved_passwords=["global1"],
            password_rules=[{"Pattern": r"backup_.*\.rar", "Passwords": ["secret1", "secret2"]}],
            archive_name="backup_2026.rar",
        )
        self.assertIn("secret1", attempts)
        self.assertIn("secret2", attempts)
        idx_s1 = attempts.index("secret1")
        idx_g1 = attempts.index("global1")
        self.assertLess(idx_s1, idx_g1)

    def test_no_match_skips_rule(self) -> None:
        attempts = build_password_attempts(
            remembered_password=None,
            saved_passwords=["global1"],
            password_rules=[{"Pattern": r"photos_.*", "Passwords": ["photo_pw"]}],
            archive_name="backup_2026.rar",
        )
        self.assertNotIn("photo_pw", attempts)

    def test_invalid_regex_ignored(self) -> None:
        attempts = build_password_attempts(
            remembered_password=None,
            saved_passwords=["x"],
            password_rules=[{"Pattern": r"[invalid", "Passwords": ["y"]}],
            archive_name="test.zip",
        )
        self.assertNotIn("y", attempts)

    def test_config_normalization(self) -> None:
        config = normalize_config({
            "PasswordRules": [
                {"Pattern": "test.*", "Passwords": ["pw1", "pw2"]},
                {"Pattern": "", "Passwords": ["empty"]},
                {"Pattern": "valid", "Passwords": []},
            ]
        })
        self.assertEqual(len(config["PasswordRules"]), 1)
        self.assertEqual(config["PasswordRules"][0]["Pattern"], "test.*")


class AutoUpdateTests(unittest.TestCase):
    def test_version_tuple_parsing(self) -> None:
        from extractorx.download import _version_tuple
        self.assertEqual(_version_tuple("2.5.0"), (2, 5, 0))
        self.assertEqual(_version_tuple("10.1"), (10, 1))
        self.assertGreater(_version_tuple("2.5.0"), _version_tuple("2.4.0"))


class WatchFolderRulesTests(unittest.TestCase):
    def test_default_empty(self) -> None:
        config = normalize_config(None)
        self.assertEqual(config["WatchFolderRules"], [])

    def test_normalizes_rules(self) -> None:
        config = normalize_config({
            "WatchFolderRules": [
                {"Folder": "C:\\Downloads", "OutputPath": "D:\\Extracted", "PostAction": "Recycle"},
                {"Folder": "", "OutputPath": "bad"},
            ]
        })
        self.assertEqual(len(config["WatchFolderRules"]), 1)
        self.assertEqual(config["WatchFolderRules"][0]["Folder"], "C:\\Downloads")


class PluginTests(unittest.TestCase):
    def test_discover_empty_dir(self) -> None:
        from extractorx.plugins import discover_plugins, plugins_dir
        directory = plugins_dir()
        if not directory.exists():
            self.assertEqual(discover_plugins(), [])

    def test_run_plugins_no_crash_on_empty(self) -> None:
        from extractorx.plugins import run_plugins
        run_plugins(Path("test.zip"), Path("output"), {})


class RepackTests(unittest.TestCase):
    def test_repack_formats_defined(self) -> None:
        from extractorx.repack import REPACK_FORMATS
        self.assertIn("zip", REPACK_FORMATS)
        self.assertIn("7z", REPACK_FORMATS)
        self.assertIn("tar", REPACK_FORMATS)

    def test_repack_rejects_unknown_format(self) -> None:
        from extractorx.repack import repack_archive
        errors: list[str] = []
        result = repack_archive(
            sevenzip_path=Path("7z"),
            source=Path("test.zip"),
            target_format="docx",
            log_cb=lambda text, level: errors.append(text),
        )
        self.assertIsNone(result)
        self.assertTrue(any("Unsupported" in e for e in errors))


class SecureDeleteTests(unittest.TestCase):
    def test_secure_delete_overwrites_content(self) -> None:
        from extractorx.postprocess import secure_delete
        with tempfile.TemporaryDirectory() as tmp:
            path = Path(tmp) / "secret.txt"
            path.write_bytes(b"sensitive data " * 1000)
            original_size = path.stat().st_size
            logs: list[str] = []
            secure_delete(path, lambda text, level: logs.append(text))
            self.assertFalse(path.exists())
            self.assertTrue(any("Securely deleted" in l for l in logs))

    def test_secure_delete_handles_large_file_chunked(self) -> None:
        from extractorx.postprocess import secure_delete
        with tempfile.TemporaryDirectory() as tmp:
            path = Path(tmp) / "large.bin"
            path.write_bytes(b"\xff" * 200_000)
            logs: list[str] = []
            secure_delete(path, lambda text, level: logs.append(text))
            self.assertFalse(path.exists())


class UniquePathBoundTests(unittest.TestCase):
    def test_unique_path_with_many_collisions(self) -> None:
        from extractorx.postprocess import _unique_path
        with tempfile.TemporaryDirectory() as tmp:
            base = Path(tmp) / "file.txt"
            base.write_text("x")
            result = _unique_path(base)
            self.assertNotEqual(result, base)
            self.assertTrue(result.name.startswith("file ("))


class ArchiveMagicStatTests(unittest.TestCase):
    def test_has_archive_magic_with_zip(self) -> None:
        import zipfile
        with tempfile.TemporaryDirectory() as tmp:
            path = Path(tmp) / "test.zip"
            with zipfile.ZipFile(path, "w") as zf:
                zf.writestr("readme.txt", "hello")
            self.assertTrue(has_archive_magic(path))

    def test_has_archive_magic_with_empty_file(self) -> None:
        with tempfile.TemporaryDirectory() as tmp:
            path = Path(tmp) / "empty.dat"
            path.write_bytes(b"")
            self.assertFalse(has_archive_magic(path))


class ThemePaletteCompletenessTests(unittest.TestCase):
    def test_all_themes_have_same_keys(self) -> None:
        from extractorx.themes import THEMES
        reference_keys = set(next(iter(THEMES.values())).keys())
        for name, palette in THEMES.items():
            self.assertEqual(set(palette.keys()), reference_keys, f"Theme '{name}' has mismatched keys")

    def test_all_themes_exist(self) -> None:
        from extractorx.themes import THEMES
        self.assertGreaterEqual(len(THEMES), 5)
        for name in ("Midnight", "Graphite", "Ocean", "White", "HighContrast"):
            self.assertIn(name, THEMES)

    def test_get_theme_fallback(self) -> None:
        from extractorx.themes import get_theme
        self.assertEqual(get_theme(None), get_theme("Midnight"))
        self.assertEqual(get_theme("nonexistent"), get_theme("Midnight"))


class GuiLogicFunctionTests(unittest.TestCase):
    def test_status_tag_working(self) -> None:
        from extractorx.ui import ExtractorXApp
        self.assertEqual(ExtractorXApp._status_tag(QueueStatus.EXTRACTING), "working")
        self.assertEqual(ExtractorXApp._status_tag(QueueStatus.TESTING), "working")

    def test_status_tag_done(self) -> None:
        from extractorx.ui import ExtractorXApp
        self.assertEqual(ExtractorXApp._status_tag(QueueStatus.DONE), "done")
        self.assertEqual(ExtractorXApp._status_tag(QueueStatus.TEST_OK), "done")

    def test_status_tag_failed(self) -> None:
        from extractorx.ui import ExtractorXApp
        self.assertEqual(ExtractorXApp._status_tag(QueueStatus.FAILED), "failed")

    def test_status_tag_queued(self) -> None:
        from extractorx.ui import ExtractorXApp
        self.assertEqual(ExtractorXApp._status_tag(QueueStatus.QUEUED), "queued")

    def test_looks_like_password_failure(self) -> None:
        from extractorx.ui import _looks_like_password_failure
        self.assertTrue(_looks_like_password_failure("Wrong password"))
        self.assertTrue(_looks_like_password_failure("file is encrypted"))
        self.assertTrue(_looks_like_password_failure("Data Error: bad password"))

    def test_looks_like_password_failure_negative(self) -> None:
        from extractorx.ui import _looks_like_password_failure
        self.assertFalse(_looks_like_password_failure("Everything is fine"))
        self.assertFalse(_looks_like_password_failure("CRC Error"))

    def test_cmd_quote(self) -> None:
        from extractorx.ui import _cmd_quote
        self.assertEqual(_cmd_quote("simple"), '"simple"')
        self.assertEqual(_cmd_quote('has "quotes"'), '"has ""quotes"""')
        self.assertEqual(_cmd_quote(""), '""')

    def test_cmd_quote_special_chars(self) -> None:
        from extractorx.ui import _cmd_quote
        result = _cmd_quote("path with spaces & special")
        self.assertTrue(result.startswith('"'))
        self.assertTrue(result.endswith('"'))


class HookExecutionTests(unittest.TestCase):
    def test_hook_runs_command_and_expands_tokens(self) -> None:
        from extractorx.hooks import run_hook
        with tempfile.TemporaryDirectory() as tmp:
            archive = Path(tmp) / "test.zip"
            archive.write_text("fake")
            output = Path(tmp) / "output"
            output.mkdir()
            marker = Path(tmp) / "hook_ran.txt"
            config = {"PostExtractCommand": f'cmd /c echo done > "{marker}"'}
            messages: list[str] = []
            run_hook("PostExtractCommand", config, archive, output, lambda msg, lvl: messages.append(msg))
            self.assertTrue(any("completed" in m or "failed" in m for m in messages))

    def test_hook_handles_nonexistent_command(self) -> None:
        from extractorx.hooks import run_hook
        with tempfile.TemporaryDirectory() as tmp:
            archive = Path(tmp) / "test.zip"
            archive.write_text("fake")
            output = Path(tmp) / "output"
            output.mkdir()
            config = {"PostExtractCommand": "nonexistent_binary_xyz123"}
            messages: list[str] = []
            run_hook("PostExtractCommand", config, archive, output, lambda msg, lvl: messages.append(msg))
            self.assertTrue(any("failed" in m.lower() for m in messages))


class EndToEndExtractionTests(unittest.TestCase):
    def test_extract_simple_zip(self) -> None:
        import zipfile
        from extractorx.extractor import ExtractionService
        from extractorx.config import DEFAULT_CONFIG
        sevenzip = _find_7zip_for_test()
        if not sevenzip:
            self.skipTest("7-Zip not available for end-to-end test")
        with tempfile.TemporaryDirectory() as tmp:
            archive = Path(tmp) / "test.zip"
            output_dir = Path(tmp) / "out"
            with zipfile.ZipFile(archive, "w") as zf:
                zf.writestr("hello.txt", "Hello, World!")
                zf.writestr("sub/nested.txt", "Nested content")
            config = dict(DEFAULT_CONFIG)
            config["OutputPath"] = str(output_dir)
            config["SmartExtract"] = "Auto"
            config["NestedExtraction"] = False
            messages: Queue = Queue()
            svc = ExtractionService(config, sevenzip, [], messages)
            item = QueueItem.from_path(archive)
            item.output_override = str(output_dir)
            svc.extract_items([item])
            svc.thread.join(timeout=30)
            self.assertEqual(item.status, QueueStatus.DONE)
            self.assertTrue((output_dir / "hello.txt").exists())
            self.assertEqual((output_dir / "hello.txt").read_text(), "Hello, World!")
            self.assertTrue((output_dir / "sub" / "nested.txt").exists())


class SanitizeFilenameTests(unittest.TestCase):
    def test_bidi_chars_stripped(self) -> None:
        from extractorx.postprocess import sanitize_extracted_filenames
        with tempfile.TemporaryDirectory() as tmp:
            bidi_name = "test‮file.txt"
            path = Path(tmp) / bidi_name
            path.write_text("content")
            messages: list[str] = []
            sanitize_extracted_filenames(Path(tmp), lambda msg, lvl: messages.append(msg))
            self.assertFalse(path.exists())
            self.assertTrue((Path(tmp) / "testfile.txt").exists())
            self.assertTrue(any("Sanitized" in m for m in messages))

    def test_reserved_name_detection(self) -> None:
        from extractorx.postprocess import _RESERVED_NAMES
        self.assertIn("con", _RESERVED_NAMES)
        self.assertIn("prn", _RESERVED_NAMES)
        self.assertIn("aux", _RESERVED_NAMES)
        self.assertIn("nul", _RESERVED_NAMES)
        self.assertIn("com1", _RESERVED_NAMES)
        self.assertIn("lpt1", _RESERVED_NAMES)
        self.assertNotIn("readme", _RESERVED_NAMES)

    def test_normal_name_unchanged(self) -> None:
        from extractorx.postprocess import sanitize_extracted_filenames
        with tempfile.TemporaryDirectory() as tmp:
            path = Path(tmp) / "normal.txt"
            path.write_text("content")
            sanitize_extracted_filenames(Path(tmp), lambda msg, lvl: None)
            self.assertTrue(path.exists())


class DllSideloadWarningTests(unittest.TestCase):
    def test_warns_when_exe_and_dll_present(self) -> None:
        from extractorx.postprocess import warn_dll_sideloading
        with tempfile.TemporaryDirectory() as tmp:
            (Path(tmp) / "app.exe").write_text("fake")
            (Path(tmp) / "payload.dll").write_text("fake")
            messages: list[str] = []
            warn_dll_sideloading(Path(tmp), lambda msg, lvl: messages.append(msg))
            self.assertTrue(any("DLL sideloading" in m for m in messages))

    def test_silent_when_exe_only(self) -> None:
        from extractorx.postprocess import warn_dll_sideloading
        with tempfile.TemporaryDirectory() as tmp:
            (Path(tmp) / "app.exe").write_text("fake")
            messages: list[str] = []
            warn_dll_sideloading(Path(tmp), lambda msg, lvl: messages.append(msg))
            self.assertEqual(len(messages), 0)

    def test_silent_when_empty(self) -> None:
        from extractorx.postprocess import warn_dll_sideloading
        with tempfile.TemporaryDirectory() as tmp:
            messages: list[str] = []
            warn_dll_sideloading(Path(tmp), lambda msg, lvl: messages.append(msg))
            self.assertEqual(len(messages), 0)


class WebhookSchemeTests(unittest.TestCase):
    def test_rejects_file_scheme(self) -> None:
        from extractorx.extractor import _send_webhook
        _send_webhook({"WebhookUrl": "file:///etc/passwd"}, {"test": True})

    def test_rejects_ftp_scheme(self) -> None:
        from extractorx.extractor import _send_webhook
        _send_webhook({"WebhookUrl": "ftp://evil.com/exfil"}, {"test": True})

    def test_accepts_https_scheme(self) -> None:
        from extractorx.extractor import _send_webhook
        _send_webhook({"WebhookUrl": "https://httpbin.org/post"}, {"test": True})

    def test_empty_url_noop(self) -> None:
        from extractorx.extractor import _send_webhook
        _send_webhook({"WebhookUrl": ""}, {"test": True})


class RepackSecurityTests(unittest.TestCase):
    def test_repack_validates_paths(self) -> None:
        import zipfile
        from extractorx.repack import repack_archive
        sevenzip = _find_7zip_for_test()
        if not sevenzip:
            self.skipTest("7-Zip not available")
        with tempfile.TemporaryDirectory() as tmp:
            archive = Path(tmp) / "test.zip"
            with zipfile.ZipFile(archive, "w") as zf:
                zf.writestr("hello.txt", "Hello")
            result = repack_archive(sevenzip, archive, "7z", output_dir=Path(tmp))
            self.assertIsNotNone(result)
            self.assertTrue(result.exists())
            self.assertTrue(result.suffix == ".7z")

    def test_repack_rejects_unknown_format(self) -> None:
        from extractorx.repack import repack_archive
        with tempfile.TemporaryDirectory() as tmp:
            archive = Path(tmp) / "test.zip"
            archive.write_text("fake")
            result = repack_archive(Path("7z"), archive, "exe")
            self.assertIsNone(result)


def _find_7zip_for_test() -> Path | None:
    from extractorx.sevenzip import find_7zip
    return find_7zip()


if __name__ == "__main__":
    unittest.main()

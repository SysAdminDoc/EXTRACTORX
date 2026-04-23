"""Core smoke tests for the Python port."""

from __future__ import annotations

import os
import tempfile
import threading
import unittest
from pathlib import Path
from queue import Queue

from extractorx.archive import archive_name, has_archive_magic, is_non_first_volume, resolve_output_path
from extractorx.config import DEFAULT_CONFIG, normalize_config
from extractorx.extractor import build_password_attempts
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

    def test_expand_processor_command_quotes_paths(self) -> None:
        template = "notepad {ArchivePath} && copy {Output} {Destination}"
        expanded = expand_processor_command(
            template,
            archive=Path(r"C:\My Stuff\demo.zip"),
            output=Path(r"D:\out\demo"),
        )
        self.assertIn(r'"C:\My Stuff\demo.zip"', expanded)
        self.assertIn(r'"D:\out\demo"', expanded)
        self.assertNotIn('""C:\\', expanded)

    def test_expand_processor_command_absorbs_wrapping_quotes(self) -> None:
        template = 'notepad "{ArchivePath}"'
        expanded = expand_processor_command(
            template,
            archive=Path(r"C:\path\file.zip"),
            output=Path(r"C:\out"),
        )
        self.assertEqual(expanded, r'notepad "C:\path\file.zip"')

    def test_expand_processor_command_escapes_embedded_quotes(self) -> None:
        expanded = expand_processor_command(
            "{ArchivePath}",
            archive=Path('C:/weird"name.zip'),
            output=Path(r"C:\out"),
        )
        self.assertTrue(expanded.startswith('"'))
        self.assertTrue(expanded.endswith('"'))
        self.assertIn('""name.zip', expanded)

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


if __name__ == "__main__":
    unittest.main()

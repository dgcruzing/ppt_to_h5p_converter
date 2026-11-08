"""Unified PPTX/PDF CLI checks; optional real mixed-folder LibreOffice conversion."""

import contextlib
import importlib
import importlib.util
import io
import json
import os
from pathlib import Path
import shutil
import subprocess
import sys
import tempfile
from types import SimpleNamespace
import unittest
from unittest import mock
import zipfile

import test_cli as pptx_tests

ROOT = Path(__file__).resolve().parents[1]
SCRIPT = ROOT / "build_script/pptx_to_h5p_converter pdf v1_4.py"
PPTX = pptx_tests.SOURCE
PDF = pptx_tests.FIXTURE / "reference/powerpoint.pdf"
spec = importlib.util.spec_from_file_location("converter_cli_v1_4", SCRIPT)
cli = importlib.util.module_from_spec(spec)
spec.loader.exec_module(cli)
engine = importlib.import_module("pptx_converter")


class UnifiedCliTests(unittest.TestCase):
    def setUp(self):
        self.temporary = tempfile.TemporaryDirectory(prefix="h5p-v1-4-tests-")
        self.addCleanup(self.temporary.cleanup)
        self.directory = Path(self.temporary.name).resolve()

    def run_cli(self, *arguments, script=SCRIPT, no_site=False):
        return subprocess.run([sys.executable, *(["-S"] if no_site else []), str(script),
                               *map(str, arguments)], capture_output=True, text=True, timeout=30)

    def test_help_without_dependencies_and_argument_errors(self):
        result = self.run_cli("--help", no_site=True)
        self.assertEqual(result.returncode, 0, result.stderr)
        for option in ("--app", "--include-hidden", "--ignore-missing-fonts", "--timeout", "--resolution"):
            self.assertIn(option, result.stdout)
        self.assertIn("ignored for PDFs", result.stdout)
        for arguments in ((PDF, "--timeout", "nan"), (PDF, "--timeout", "0"),
                          (PDF, "--resolution", "medium"), (self.directory,)):
            result = self.run_cli(*arguments)
            self.assertEqual(result.returncode, 2, result.stderr)
            self.assertNotIn("Traceback", result.stderr)
        result = self.run_cli(PDF, no_site=True)
        self.assertEqual(result.returncode, 1)
        self.assertIn("pip install -r requirements.txt", result.stderr)
        self.assertNotIn("Traceback", result.stderr)

    def test_mixed_discovery_order_defaults_and_unsupported_ppt(self):
        for name in ("z.PDF", "A.PPTX", "b.pdf", "~$lock.pptx", "legacy.ppt", "ignored.txt"):
            (self.directory / name).touch()
        nested = self.directory / "nested.pdf"
        nested.mkdir()
        (nested / "skip.pptx").touch()
        jobs = cli.build_jobs(self.directory)
        self.assertEqual([source.name for source, _ in jobs], ["A.PPTX", "b.pdf", "z.PDF"])
        self.assertEqual([output.name for _, output in jobs], ["A.h5p", "b.h5p", "z.h5p"])
        for source in (PPTX, PDF):
            self.assertEqual(cli.build_jobs(source), [(source, source.with_suffix(".h5p"))])
        with self.assertRaises(ValueError):
            cli.build_jobs(self.directory / "legacy.ppt")
        result = self.run_cli(self.directory / "legacy.ppt")
        self.assertEqual(result.returncode, 2)

    def test_cross_format_name_collisions_fail_before_output_creation(self):
        (self.directory / "course.pptx").touch()
        (self.directory / "course.pdf").touch()
        output = self.directory / "converted"
        result = self.run_cli(self.directory, "-o", output, "--force")
        self.assertEqual(result.returncode, 2, result.stderr)
        self.assertIn("same output", result.stderr)
        self.assertFalse(output.exists())
        self.assertEqual(result.stdout, "")

    def test_pdf_only_needs_no_pptx_or_font_modules_and_ignores_presentation_options(self):
        from PIL import Image

        # Distribute only the unified script and common helper: PDF conversion must work.
        script = self.directory / SCRIPT.name
        shutil.copyfile(SCRIPT, script)
        shutil.copyfile(SCRIPT.parent / "h5p_converter_common.py", self.directory / "h5p_converter_common.py")
        output = self.directory / "pdf/course.h5p"
        result = self.run_cli(PDF, "-o", output, "--resolution", "low", "--app", "powerpoint",
                              "--libreoffice-path", self.directory / "missing", "--include-hidden",
                              "--ignore-missing-fonts", "--timeout", "0.01", script=script)
        self.assertEqual(result.returncode, 0, result.stderr)
        self.assertEqual(result.stdout.strip(), str(output))
        with zipfile.ZipFile(output) as archive:
            self.assertEqual(len(json.loads(archive.read("content/content.json"))["presentation"]["slides"]), 7)
            with Image.open(io.BytesIO(archive.read("content/images/slide_1.png"))) as image:
                self.assertEqual(image.size, (1280, 720))

    def test_real_mixed_packaging_and_presentation_flags(self):
        source = self.directory / "Mixed documents"
        source.mkdir()
        pptx_tests.variant(source / "A hidden.PPTX", hidden=(2, 5))
        shutil.copyfile(PDF, source / "B reference.PDF")
        output = self.directory / "Converted"
        # Only the Office export is mocked; both formats are really rasterized and packaged.
        with mock.patch.object(engine, "find_libreoffice", return_value=Path("/fake/soffice")), \
                mock.patch.object(engine, "missing_fonts", side_effect=AssertionError("must skip")), \
                mock.patch.object(engine, "export_pdf", return_value=PDF) as export, \
                contextlib.redirect_stdout(io.StringIO()) as stdout, \
                contextlib.redirect_stderr(io.StringIO()) as stderr:
            result = cli.main([str(source), "-o", str(output), "--resolution", "low", "--app", "libreoffice",
                               "--include-hidden", "--ignore-missing-fonts", "--timeout", "91"])
        self.assertEqual(result, 0, stderr.getvalue())
        self.assertIn("2 succeeded, 0 failed", stderr.getvalue())
        self.assertEqual(stdout.getvalue().splitlines(), [str(output / "A hidden.h5p"), str(output / "B reference.h5p")])
        self.assertEqual(export.call_args.args[3:5], (True, 91.0))
        with zipfile.ZipFile(output / "A hidden.h5p") as presentation, \
                zipfile.ZipFile(output / "B reference.h5p") as document:
            for archive, title in ((presentation, "A hidden"), (document, "B reference")):
                self.assertIsNone(archive.testzip())
                self.assertEqual(json.loads(archive.read("h5p.json"))["title"], title)
                slides = json.loads(archive.read("content/content.json"))["presentation"]["slides"]
                self.assertEqual(len(slides), 7)
            for number in range(1, 8):
                name = f"content/images/slide_{number}.png"
                self.assertEqual(presentation.read(name), document.read(name))

    def test_missing_office_fails_presentations_but_converts_later_pdf(self):
        for name in ("A.pptx", "B.PPTX"):
            shutil.copyfile(PPTX, self.directory / name)
        shutil.copyfile(PDF, self.directory / "C.pdf")
        output = self.directory / "converted"
        with mock.patch.object(engine, "find_libreoffice", side_effect=cli.ConversionError("Office missing")) as find, \
                contextlib.redirect_stdout(io.StringIO()) as stdout, \
                contextlib.redirect_stderr(io.StringIO()) as stderr:
            result = cli.main([str(self.directory), "-o", str(output), "--app", "libreoffice", "--resolution", "low"])
        self.assertEqual(result, 1)
        self.assertIn("1 succeeded, 2 failed", stderr.getvalue())
        self.assertIn("Failed A.pptx: Office missing", stderr.getvalue())
        self.assertIn("Failed B.PPTX: Office missing", stderr.getvalue())
        self.assertNotIn("Traceback", stderr.getvalue())
        self.assertEqual(stdout.getvalue().strip(), str(output / "C.h5p"))
        self.assertTrue(zipfile.is_zipfile(output / "C.h5p"))
        find.assert_called_once()

    def test_pptx_font_check_failure_does_not_block_pdf(self):
        shutil.copyfile(PPTX, self.directory / "A.pptx")
        shutil.copyfile(PDF, self.directory / "B.pdf")
        with mock.patch.object(engine, "find_libreoffice", return_value=Path("/fake/soffice")), \
                mock.patch.object(engine, "missing_fonts", return_value=[SimpleNamespace(family="Missing", style="Bold")]) as fonts, \
                mock.patch.object(engine, "export_pdf") as export, \
                contextlib.redirect_stdout(io.StringIO()), contextlib.redirect_stderr(io.StringIO()) as stderr:
            result = cli.main([str(self.directory), "--app", "libreoffice", "--resolution", "low"])
        self.assertEqual(result, 1)
        self.assertIn("Missing — Bold", stderr.getvalue())
        self.assertIn("1 succeeded, 1 failed", stderr.getvalue())
        self.assertFalse((self.directory / "A.h5p").exists())
        self.assertTrue(zipfile.is_zipfile(self.directory / "B.h5p"))
        fonts.assert_called_once()
        export.assert_not_called()

    def test_powerpoint_dispatch_uses_existing_engine(self):
        args = cli.create_parser().parse_args([str(PPTX), "--app", "powerpoint"])
        output = self.directory / "course.h5p"
        with mock.patch.object(cli, "require_module") as require, \
                mock.patch.object(engine, "convert") as convert:
            cli.convert(PPTX, output, args, app="powerpoint")
        self.assertEqual([call.args[0] for call in require.call_args_list],
                         ["pythoncom", "win32com.client", "win32process"])
        convert.assert_called_once_with(PPTX, output, "powerpoint", None, args)

    def test_powerpoint_worker_uses_the_v1_4_backend_module(self):
        from PIL import Image

        args = cli.create_parser().parse_args([str(PPTX), "--resolution", "low"])
        numbers = list(range(1, 8))

        def worker_response(command, *_):
            request = json.loads(Path(command[3]).read_text())
            for number in numbers:
                shutil.copyfile(pptx_tests.REFERENCE_IMAGE, Path(request["directory"]) / f"slide_{number}.png")
            Path(request["result_file"]).write_text(json.dumps(numbers))

        with mock.patch.object(engine, "run_export", side_effect=worker_response) as run:
            images = engine.render_powerpoint(PPTX, self.directory, (1280, 720), numbers, args)
        command = run.call_args.args[0]
        self.assertEqual(Path(command[1]), Path(engine.__file__).resolve())
        self.assertEqual(command[2], "--_powerpoint-worker")
        self.assertEqual([number for number, _ in images], numbers)
        with Image.open(images[0][1]) as image:
            self.assertEqual(image.size, (1280, 720))

    def test_failed_pdf_does_not_block_later_pptx(self):
        (self.directory / "A.pdf").write_bytes(b"invalid PDF")
        shutil.copyfile(PPTX, self.directory / "B.pptx")
        with mock.patch.object(engine, "find_libreoffice", return_value=Path("/fake/soffice")), \
                mock.patch.object(engine, "missing_fonts", return_value=[]), \
                mock.patch.object(engine, "export_pdf", return_value=PDF), \
                contextlib.redirect_stdout(io.StringIO()), contextlib.redirect_stderr(io.StringIO()) as stderr:
            result = cli.main([str(self.directory), "--app", "libreoffice", "--resolution", "low"])
        self.assertEqual(result, 1)
        self.assertIn("1 succeeded, 1 failed", stderr.getvalue())
        self.assertTrue(zipfile.is_zipfile(self.directory / "B.h5p"))
        self.assertFalse((self.directory / "A.h5p").exists())

    def test_output_protection_force_and_packaging_failure(self):
        output = self.directory / "course.h5p"
        output.write_bytes(b"original")
        args = cli.create_parser().parse_args([str(PDF), "--resolution", "low"])
        with mock.patch.object(cli, "render_pdf") as render:
            with self.assertRaisesRegex(cli.ConversionError, "Output exists"):
                cli.convert(PDF, output, args)
            render.assert_not_called()
        args.force = True
        with mock.patch.object(cli, "write_h5p", side_effect=OSError("disk full")) as write:
            with self.assertRaisesRegex(OSError, "disk full"):
                cli.convert(PDF, output, args)
        self.assertEqual(output.read_bytes(), b"original")
        self.assertEqual(list(self.directory.iterdir()), [output])
        self.assertFalse(write.call_args.args[2][0][1].parent.exists())
        cli.convert(PDF, output, args)
        self.assertTrue(zipfile.is_zipfile(output))

    def test_interrupt_returns_130_and_cleans_up(self):
        output = self.directory / "course.h5p"
        with mock.patch.object(cli, "write_h5p", side_effect=KeyboardInterrupt) as write, \
                contextlib.redirect_stderr(io.StringIO()):
            result = cli.main([str(PDF), "-o", str(output), "--resolution", "low"])
        self.assertEqual(result, 130)
        self.assertFalse(output.exists())
        self.assertEqual(list(self.directory.iterdir()), [])
        self.assertFalse(write.call_args.args[2][0][1].parent.exists())


@unittest.skipUnless(os.environ.get("H5P_RUN_INTEGRATION") == "1", "Set H5P_RUN_INTEGRATION=1")
class MixedFolderIntegrationTests(unittest.TestCase):
    def test_actual_libreoffice_and_pdf_in_one_folder(self):
        from PIL import Image

        with tempfile.TemporaryDirectory(prefix="h5p-v1-4-integration-") as temporary:
            directory = Path(temporary).resolve()
            # Verify distribution without the legacy CLI or standalone PDF entry point.
            scripts = directory / "scripts"
            scripts.mkdir()
            for name in (SCRIPT.name, "h5p_converter_common.py", "pptx_converter.py",
                         "pptx_to_h5p_converter cli v1_3 verify fonts.py"):
                shutil.copyfile(SCRIPT.parent / name, scripts / name)
            source = directory / "Mixed documents"
            source.mkdir()
            (source / "A broken.pdf").write_bytes(b"not a PDF")
            pptx_tests.variant(source / "B hidden.PPTX", hidden=(2, 5))
            shutil.copyfile(PDF, source / "C référence.PDF")
            (source / "legacy.ppt").touch()
            (source / "~$locked.pptx").touch()
            (source / "nested").mkdir()
            shutil.copyfile(PDF, source / "nested/skip.pdf")
            output = directory / "Converted"
            command = [sys.executable, str(scripts / SCRIPT.name), str(source), "-o", str(output), "--app", "libreoffice",
                       "--resolution", "low", "--ignore-missing-fonts", "--timeout", "60"]
            result = subprocess.run(command, capture_output=True, text=True, timeout=90)
            self.assertEqual(result.returncode, 1, result.stderr)
            self.assertIn("2 succeeded, 1 failed", result.stderr)
            self.assertEqual(result.stdout.splitlines(), [str(output / "B hidden.h5p"), str(output / "C référence.h5p")])
            for filename, numbers in (("B hidden.h5p", [1, 3, 4, 6, 7]), ("C référence.h5p", list(range(1, 8)))):
                with zipfile.ZipFile(output / filename) as archive:
                    self.assertIsNone(archive.testzip())
                    slides = json.loads(archive.read("content/content.json"))["presentation"]["slides"]
                    self.assertEqual([slide["elements"][0]["action"]["metadata"]["title"] for slide in slides],
                                     [f"Slide {number}" for number in numbers])
                    for position in range(1, len(slides) + 1):
                        with Image.open(io.BytesIO(archive.read(f"content/images/slide_{position}.png"))) as image:
                            self.assertEqual(image.size, (1280, 720))


if __name__ == "__main__":
    unittest.main()

"""Run with unittest; set H5P_RUN_INTEGRATION=1 for real LibreOffice exports."""

import contextlib
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
import xml.etree.ElementTree as ET
import zipfile

ROOT = Path(__file__).resolve().parents[1]
SCRIPT = ROOT / "build_script/pptx_to_h5p_converter cli v1_3.py"
sys.path.insert(0, str(SCRIPT.parent))
FIXTURE = ROOT / "tests/fixtures/introduction_to_h5p"
SOURCE = FIXTURE / "source.pptx"
REFERENCE_IMAGE = FIXTURE / "reference/powerpoint_images/slide_01.png"
spec = importlib.util.spec_from_file_location("converter_cli", SCRIPT)
cli = importlib.util.module_from_spec(spec)
spec.loader.exec_module(cli)


def variant(destination, hidden=(), four_by_three=False, reverse=False):
    """Derive decks without changing the user's reference presentation."""
    with zipfile.ZipFile(SOURCE) as source, zipfile.ZipFile(destination, "w") as output:
        for item in source.infolist():
            data = source.read(item.filename)
            if item.filename in {f"ppt/slides/slide{n}.xml" for n in hidden}:
                root = ET.fromstring(data)
                root.set("show", "0")
                data = ET.tostring(root, encoding="utf-8", xml_declaration=True)
            elif item.filename == "ppt/presentation.xml" and (four_by_three or reverse):
                root = ET.fromstring(data)
                if four_by_three:
                    size = root.find(f"{{{cli.PRESENTATION_NS}}}sldSz")
                    size.set("cx", "9144000")
                    size.set("cy", "6858000")
                if reverse:
                    slides = root.find(f"{{{cli.PRESENTATION_NS}}}sldIdLst")
                    slides[:] = list(reversed(slides))
                data = ET.tostring(root, encoding="utf-8", xml_declaration=True)
            output.writestr(item, data)


class CliTests(unittest.TestCase):
    def setUp(self):
        self.temporary = tempfile.TemporaryDirectory()
        self.addCleanup(self.temporary.cleanup)
        self.directory = Path(self.temporary.name)
        # These tests exercise conversion mechanics; font preflight has its own tests.
        patch = mock.patch.object(cli, "missing_fonts", return_value=[])
        patch.start()
        self.addCleanup(patch.stop)

    def test_platform_selection(self):
        for platform, expected in (("win32", "powerpoint"), ("darwin", "libreoffice"),
                                   ("linux", "libreoffice")):
            self.assertEqual(cli.select_app("auto", platform), expected)
        self.assertEqual(cli.select_app("libreoffice", "win32"), "libreoffice")
        self.assertEqual(cli.select_app("powerpoint", "win32"), "powerpoint")
        with self.assertRaisesRegex(ValueError, "requires Windows"):
            cli.select_app("powerpoint", "darwin")

    def test_help_without_site_packages(self):
        result = subprocess.run([sys.executable, "-S", str(SCRIPT), "--help"], capture_output=True)
        self.assertEqual(result.returncode, 0, result.stderr)
        self.assertIn(b"--libreoffice-path", result.stdout)

    def test_missing_dependency_without_traceback(self):
        result = subprocess.run([sys.executable, "-S", str(SCRIPT), str(SOURCE)], capture_output=True)
        self.assertEqual(result.returncode, 1)
        self.assertIn(b"pip install -r requirements.txt", result.stderr)
        self.assertNotIn(b"Traceback", result.stderr)

    def test_invalid_arguments(self):
        for arguments in (["--timeout", "0"], ["--timeout", "-1"], ["--timeout", "nan"],
                          ["--timeout", "inf"], ["--timeout", "bad"], ["--resolution", "medium"]):
            with self.subTest(arguments=arguments), contextlib.redirect_stderr(io.StringIO()):
                with self.assertRaises(SystemExit) as error:
                    cli.main([str(SOURCE), *arguments])
                self.assertEqual(error.exception.code, 2)
        self.assertEqual(cli.create_parser().parse_args([str(SOURCE)]).timeout, 30)
        self.assertEqual(cli.create_parser().parse_args([str(SOURCE), "--timeout", "0.2"]).timeout, 0.2)

    def test_invalid_inputs_and_outputs(self):
        for source, output in ((self.directory / "missing.pptx", None), (self.directory, None),
                               (SOURCE, self.directory), (SOURCE, self.directory / "bad.zip")):
            with self.subTest(source=source, output=output), self.assertRaises(ValueError):
                cli.build_jobs(source, output)

    def test_folder_filter_order_and_default_output(self):
        for name in ("z.PPTX", "A.pptx", "~$lock.pptx", "ignored.pdf"):
            (self.directory / name).touch()
        (self.directory / "nested").mkdir()
        (self.directory / "nested/skip.pptx").touch()
        jobs = cli.build_jobs(self.directory)
        self.assertEqual([p.name for p, _ in jobs], ["A.pptx", "z.PPTX"])
        self.assertEqual([p.name for _, p in jobs], ["A.h5p", "z.h5p"])
        self.assertEqual(cli.build_jobs(SOURCE)[0][1], SOURCE.with_suffix(".h5p"))

    def test_folder_output_conflicts(self):
        first = self.directory / "same.pptx"
        second = self.directory / "same.PPTX"
        first.touch()
        second.touch()
        # Simulate a case-sensitive source listing even on a case-insensitive Mac.
        with mock.patch.object(Path, "iterdir", return_value=iter((first, second))):
            with self.assertRaisesRegex(ValueError, "same output"):
                cli.build_jobs(self.directory)

    def test_slide_relationship_order_and_hidden_filter(self):
        path = self.directory / "hidden.pptx"
        variant(path, hidden=(2, 5))
        self.assertEqual(cli.eligible_slides(path), [1, 3, 4, 6, 7])
        self.assertEqual(cli.eligible_slides(path, True), list(range(1, 8)))
        variant(path, hidden=(2,), reverse=True)
        self.assertEqual(cli.eligible_slides(path), [1, 2, 3, 4, 5, 7])

    def test_all_hidden_and_corrupt_decks(self):
        path = self.directory / "all hidden.pptx"
        variant(path, hidden=range(1, 8))
        with self.assertRaisesRegex(cli.ConversionError, "No eligible slides"):
            cli.eligible_slides(path)
        self.assertEqual(len(cli.eligible_slides(path, True)), 7)
        path.write_bytes(b"not a pptx")
        with self.assertRaisesRegex(cli.ConversionError, "Invalid or unreadable"):
            cli.eligible_slides(path)

    def test_discovery_and_invalid_explicit_path(self):
        with mock.patch.object(cli.shutil, "which", return_value="/fake/soffice"):
            self.assertEqual(cli.find_libreoffice(), Path("/fake/soffice"))
            with self.assertRaisesRegex(cli.ConversionError, "Invalid LibreOffice"):
                cli.find_libreoffice(self.directory / "missing")
        with mock.patch.object(cli.shutil, "which", return_value=None), \
                mock.patch.object(Path, "is_file", return_value=False):
            with self.assertRaisesRegex(cli.ConversionError, "Install LibreOffice"):
                cli.find_libreoffice()

    def test_pdf_export_options_and_missing_output(self):
        for hidden in (False, True):
            with mock.patch.object(cli, "run_export", return_value="export failed") as run:
                with self.assertRaisesRegex(cli.ConversionError, "did not produce a PDF"):
                    cli.export_pdf(SOURCE, self.directory, Path("soffice"), hidden, 90)
                command, timeout, *_ = run.call_args.args
                self.assertEqual(timeout, 90)
                self.assertIn(f"-env:UserInstallation={(self.directory / 'profile').as_uri()}", command)
                options = json.loads(command[command.index("--convert-to") + 1].split(":", 2)[2])
                self.assertEqual(options["ExportHiddenSlides"]["value"], hidden)
                self.assertFalse(options["ExportNotes"]["value"])
                self.assertFalse(options["ExportNotesPages"]["value"])

    def test_pdf_comments_are_omitted_without_removing_slide_content(self):
        import pypdfium2 as pdfium
        from PIL import Image

        def stream(content, attributes=b""):
            return (b"<< " + attributes + f" /Length {len(content)} >>\nstream\n".encode()
                    + content + b"\nendstream")

        # A blue slide object and a separate yellow review-comment annotation.
        objects = [
            b"<< /Type /Catalog /Pages 2 0 R >>",
            b"<< /Type /Pages /Kids [3 0 R] /Count 1 >>",
            b"<< /Type /Page /Parent 2 0 R /MediaBox [0 0 100 100] "
            b"/Resources << >> /Contents 4 0 R /Annots [5 0 R] >>",
            stream(b"0 0 1 rg 10 10 20 20 re f"),
            b"<< /Type /Annot /Subtype /Text /Rect [40 40 60 60] "
            b"/Contents (Review comment) /AP << /N 6 0 R >> >>",
            stream(b"1 1 0 rg 0 0 20 20 re f",
                   b"/Type /XObject /Subtype /Form /BBox [0 0 20 20] /Resources << >>"),
        ]
        data = bytearray(b"%PDF-1.4\n")
        offsets = [0]
        for number, obj in enumerate(objects, 1):
            offsets.append(len(data))
            data.extend(f"{number} 0 obj\n".encode() + obj + b"\nendobj\n")
        xref = len(data)
        data.extend(f"xref\n0 {len(offsets)}\n0000000000 65535 f \n".encode())
        for offset in offsets[1:]:
            data.extend(f"{offset:010d} 00000 n \n".encode())
        data.extend(f"trailer\n<< /Size {len(offsets)} /Root 1 0 R >>\n"
                    f"startxref\n{xref}\n%%EOF\n".encode())
        pdf = self.directory / "comment.pdf"
        pdf.write_bytes(data)

        # Ensure the fixture really shows the comment with default rendering.
        with pdfium.PdfDocument(pdf) as document, contextlib.closing(document[0]) as page:
            with contextlib.closing(page.render()) as bitmap, bitmap.to_pil() as image:
                self.assertEqual(image.getpixel((50, 50)), (255, 255, 0))

        images = cli.render_pdf(pdf, self.directory, (100, 100), [12])
        self.assertEqual([number for number, _ in images], [12])
        with Image.open(images[0][1]) as image:
            self.assertEqual(image.getpixel((50, 50)), (255, 255, 255))
            self.assertEqual(image.getpixel((20, 80)), (0, 0, 255))

    def test_real_subprocess_timeout_and_failure(self):
        with self.assertRaisesRegex(cli.ConversionError, "exceeded 0.1 seconds"):
            cli.run_export([sys.executable, "-c", "import time; time.sleep(10)"], 0.1, "Test")
        with self.assertRaisesRegex(cli.ConversionError, "exit 3.*intentional"):
            cli.run_export([sys.executable, "-c", "import sys; sys.stderr.write('intentional'); sys.exit(3)"],
                           5, "Test")

    def test_interrupt_stops_export(self):
        process = mock.Mock()
        process.communicate.side_effect = KeyboardInterrupt
        with mock.patch.object(cli.subprocess, "Popen", return_value=process), \
                mock.patch.object(cli, "stop_process") as stop:
            with self.assertRaises(KeyboardInterrupt):
                cli.run_export(["fake"], 30, "Test")
            stop.assert_called_once_with(process)

    def test_normalization_padding_and_alpha(self):
        from PIL import Image
        destination = self.directory / "padded.png"
        with Image.new("RGBA", (400, 300), (255, 0, 0, 128)) as source:
            cli.normalize_image(source, (1280, 720), destination)
        with Image.open(destination) as image:
            self.assertEqual(image.size, (1280, 720))
            self.assertEqual(image.getpixel((0, 360)), (255, 255, 255))
            self.assertEqual(image.getpixel((640, 360)), (255, 127, 127))

    def test_archive_structure_and_image_references(self):
        from PIL import Image
        archive_path = self.directory / "output.h5p"
        cli.write_h5p(SOURCE, archive_path, [(1, REFERENCE_IMAGE), (3, REFERENCE_IMAGE)], (720, 405))
        with zipfile.ZipFile(archive_path) as archive:
            self.assertIsNone(archive.testzip())
            self.assertEqual(set(archive.namelist()), {
                "h5p.json", "content/content.json", "content/images/slide_1.png", "content/images/slide_2.png"})
            metadata = json.loads(archive.read("h5p.json"))
            self.assertEqual(metadata["title"], SOURCE.stem)
            self.assertIn({"machineName": "H5P.Image", "majorVersion": 1, "minorVersion": 1},
                          metadata["preloadedDependencies"])
            slides = json.loads(archive.read("content/content.json"))["presentation"]["slides"]
            ids = []
            for number, slide in zip((1, 3), slides):
                action = slide["elements"][0]["action"]
                self.assertEqual(action["metadata"]["title"], f"Slide {number}")
                file = action["params"]["file"]
                with Image.open(io.BytesIO(archive.read("content/" + file["path"]))) as image:
                    self.assertEqual(image.size, (file["width"], file["height"]))
                ids.append(action["subContentId"])
            self.assertEqual(len(set(ids)), 2)

    def test_output_protection_cleanup_and_atomic_publish(self):
        args = cli.create_parser().parse_args([str(SOURCE), "--resolution", "low"])
        destination = self.directory / "result.h5p"
        destination.write_bytes(b"original")
        with mock.patch.object(cli, "export_pdf") as export:
            with self.assertRaisesRegex(cli.ConversionError, "Output exists"):
                cli.convert(SOURCE, destination, "libreoffice", None, args)
            export.assert_not_called()
        args.force = True
        with mock.patch.object(cli, "export_pdf") as export, \
                mock.patch.object(cli, "render_pdf", return_value=[(1, REFERENCE_IMAGE)]), \
                mock.patch.object(cli, "write_h5p", side_effect=RuntimeError("packaging failed")):
            with self.assertRaisesRegex(RuntimeError, "packaging failed"):
                cli.convert(SOURCE, destination, "libreoffice", None, args)
            self.assertFalse(export.call_args.args[1].exists())
            self.assertEqual(destination.read_bytes(), b"original")
            self.assertEqual(list(self.directory.iterdir()), [destination])
        args.force = False
        destination.unlink()
        with mock.patch.object(cli, "export_pdf"), \
                mock.patch.object(cli, "render_pdf", return_value=[(1, REFERENCE_IMAGE)]):
            cli.convert(SOURCE, destination, "libreoffice", None, args)
        self.assertTrue(zipfile.is_zipfile(destination))
        self.assertEqual(list(self.directory.iterdir()), [destination])
        args.force = True
        destination.write_bytes(b"replace me")
        with mock.patch.object(cli, "export_pdf"), \
                mock.patch.object(cli, "render_pdf", return_value=[(1, REFERENCE_IMAGE)]):
            cli.convert(SOURCE, destination, "libreoffice", None, args)
        self.assertTrue(zipfile.is_zipfile(destination))

    def test_output_created_during_export_is_protected(self):
        args = cli.create_parser().parse_args([str(SOURCE)])
        destination = self.directory / "raced.h5p"

        def competing_writer(*_):
            destination.write_bytes(b"created by someone else")

        with mock.patch.object(cli, "export_pdf"), \
                mock.patch.object(cli, "render_pdf", return_value=[(1, REFERENCE_IMAGE)]), \
                mock.patch.object(cli, "write_h5p", side_effect=competing_writer):
            with self.assertRaisesRegex(cli.ConversionError, "Output exists"):
                cli.convert(SOURCE, destination, "libreoffice", None, args)
        self.assertEqual(destination.read_bytes(), b"created by someone else")
        self.assertEqual(list(self.directory.iterdir()), [destination])

    def test_all_hidden_conversion_does_not_launch_app_or_write_output(self):
        source = self.directory / "all hidden.pptx"
        variant(source, hidden=range(1, 8))
        args = cli.create_parser().parse_args([str(source)])
        destination = source.with_suffix(".h5p")
        with mock.patch.object(cli, "export_pdf") as export:
            with self.assertRaisesRegex(cli.ConversionError, "No eligible slides"):
                cli.convert(source, destination, "libreoffice", None, args)
            export.assert_not_called()
        self.assertFalse(destination.exists())

    def test_folder_continues_after_failures_and_interrupt_code(self):
        for name in ("a.pptx", "b.pptx", "c.pptx"):
            (self.directory / name).touch()
        with mock.patch.object(cli, "require_module"), mock.patch.object(cli, "find_libreoffice"), \
                mock.patch.object(cli, "convert", side_effect=[ValueError("bad deck"), None, None]) as convert, \
                contextlib.redirect_stdout(io.StringIO()) as stdout, \
                contextlib.redirect_stderr(io.StringIO()) as stderr:
            self.assertEqual(cli.main([str(self.directory), "--app", "libreoffice"]), 1)
            self.assertEqual(convert.call_count, 3)
            self.assertIn("2 succeeded, 1 failed", stderr.getvalue())
            self.assertIn("b.h5p", stdout.getvalue())
        with mock.patch.object(cli, "require_module"), mock.patch.object(cli, "find_libreoffice"), \
                mock.patch.object(cli, "convert", side_effect=KeyboardInterrupt), \
                contextlib.redirect_stderr(io.StringIO()):
            self.assertEqual(cli.main([str(SOURCE), "--app", "libreoffice"]), 130)

    def test_powerpoint_filters_and_preserves_existing_session(self):
        for existing in ({77}, set()):
            with self.subTest(existing=existing):
                slides = [mock.Mock(SlideShowTransition=SimpleNamespace(Hidden=hidden))
                          for hidden in (0, -1, 0)]
                presentation = mock.Mock(PageSetup=SimpleNamespace(SlideWidth=400, SlideHeight=300))
                presentation.Slides.Count = len(slides)
                presentation.Slides.Item.side_effect = lambda index: slides[index - 1]
                app = mock.Mock(HWND=99)
                app.Presentations.Open.return_value = presentation
                modules = {"pythoncom": mock.Mock(), "win32com.client": mock.Mock(),
                           "win32process": mock.Mock()}
                modules["win32com.client"].DispatchEx.return_value = app
                modules["win32process"].GetWindowThreadProcessId.return_value = (1, 77)
                ownership = self.directory / "owner.json"
                with mock.patch.object(cli, "require_module", side_effect=lambda name, _: modules[name]), \
                        mock.patch.object(cli, "powerpoint_pids", return_value=existing):
                    images = cli.export_powerpoint(SOURCE, self.directory, (1280, 720), False, ownership)
                self.assertEqual([n for n, _ in images], [1, 3])
                slides[1].Export.assert_not_called()
                slides[0].Export.assert_called_once_with(str(self.directory / "slide_1.png"), "PNG", 960, 720)
                presentation.Close.assert_called_once()
                self.assertEqual(app.Quit.call_count, 0 if existing else 1)
                self.assertEqual(ownership.exists(), not bool(existing))
                modules["pythoncom"].CoUninitialize.assert_called_once()

    def test_powerpoint_includes_hidden_and_closes_after_failure(self):
        slide = mock.Mock(SlideShowTransition=SimpleNamespace(Hidden=-1))
        presentation = mock.Mock(PageSetup=SimpleNamespace(SlideWidth=1600, SlideHeight=900))
        presentation.Slides.Count = 1
        presentation.Slides.Item.return_value = slide
        app = mock.Mock(HWND=99)
        app.Presentations.Open.return_value = presentation
        modules = {"pythoncom": mock.Mock(), "win32com.client": mock.Mock(), "win32process": mock.Mock()}
        modules["win32com.client"].DispatchEx.return_value = app
        modules["win32process"].GetWindowThreadProcessId.return_value = (1, 77)
        with mock.patch.object(cli, "require_module", side_effect=lambda name, _: modules[name]), \
                mock.patch.object(cli, "powerpoint_pids", return_value=set()):
            images = cli.export_powerpoint(SOURCE, self.directory, (1280, 720), True,
                                          self.directory / "owner.json")
            self.assertEqual(len(images), 1)
            slide.Export.side_effect = RuntimeError("export failure")
            with self.assertRaisesRegex(RuntimeError, "export failure"):
                cli.export_powerpoint(SOURCE, self.directory, (1280, 720), True,
                                      self.directory / "owner.json")
        self.assertEqual(presentation.Close.call_count, 2)
        self.assertEqual(app.Quit.call_count, 2)


@unittest.skipUnless(os.environ.get("H5P_RUN_INTEGRATION") == "1", "Set H5P_RUN_INTEGRATION=1")
class LibreOfficeIntegrationTests(unittest.TestCase):
    def test_reference_deck_and_folder_variants(self):
        from PIL import Image
        with tempfile.TemporaryDirectory(prefix="h5p-tests-") as temporary:
            directory = Path(temporary)
            source = directory / "Présentations with spaces"
            source.mkdir()
            shutil.copyfile(SOURCE, source / "A reference.pptx")
            variant(source / "B hidden.pptx", hidden=(2, 5))
            variant(source / "C four by three.PPTX", four_by_three=True)
            (source / "D corrupt.pptx").write_bytes(b"corrupt")
            (source / "~$lock.pptx").touch()
            (source / "ignored.pdf").touch()
            (source / "nested").mkdir()
            shutil.copyfile(SOURCE, source / "nested/skip.pptx")
            output = directory / "Converted courses"
            command = [sys.executable, str(SCRIPT), str(source), "-o", str(output),
                       "--app", "libreoffice", "--resolution", "low", "--ignore-missing-fonts"]
            result = subprocess.run(command, capture_output=True, timeout=150)
            self.assertEqual(result.returncode, 1, result.stderr.decode())
            self.assertIn(b"3 succeeded, 1 failed", result.stderr)
            self.assertEqual(len(list(output.glob("*.h5p"))), 3)
            # Same renderer/resolution: retained slide images must match the originals.
            with zipfile.ZipFile(output / "A reference.h5p") as original, \
                    zipfile.ZipFile(output / "B hidden.h5p") as filtered:
                for position, number in enumerate((1, 3, 4, 6, 7), 1):
                    self.assertEqual(filtered.read(f"content/images/slide_{position}.png"),
                                     original.read(f"content/images/slide_{number}.png"))
            for filename, numbers in (("A reference.h5p", list(range(1, 8))),
                                      ("B hidden.h5p", [1, 3, 4, 6, 7]),
                                      ("C four by three.h5p", list(range(1, 8)))):
                with zipfile.ZipFile(output / filename) as archive:
                    slides = json.loads(archive.read("content/content.json"))["presentation"]["slides"]
                    self.assertEqual([s["elements"][0]["action"]["metadata"]["title"] for s in slides],
                                     [f"Slide {n}" for n in numbers])
                    for index in range(1, len(slides) + 1):
                        with Image.open(io.BytesIO(archive.read(f"content/images/slide_{index}.png"))) as image:
                            self.assertEqual(image.size, (1280, 720))
                            if filename.startswith("C"):
                                self.assertEqual(image.getpixel((0, 360)), (255, 255, 255))
            hidden_output = directory / "including hidden.h5p"
            result = subprocess.run([sys.executable, str(SCRIPT), str(source / "B hidden.pptx"),
                                     "-o", str(hidden_output), "--include-hidden", "--timeout", "60",
                                     "--ignore-missing-fonts"],
                                    capture_output=True, timeout=90)
            self.assertEqual(result.returncode, 0, result.stderr.decode())
            with zipfile.ZipFile(hidden_output) as archive:
                slides = json.loads(archive.read("content/content.json"))["presentation"]["slides"]
                self.assertEqual(len(slides), 7)
                with Image.open(io.BytesIO(archive.read("content/images/slide_1.png"))) as image:
                    self.assertEqual(image.size, (1920, 1080))


if __name__ == "__main__":
    unittest.main()

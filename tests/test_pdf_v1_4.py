"""PDF input checks for the unified v1.4 CLI, without an Office installation."""

import contextlib
import importlib.util
import io
import json
from pathlib import Path
import shutil
import subprocess
import sys
import tempfile
import unittest
from unittest import mock
import zipfile

ROOT = Path(__file__).resolve().parents[1]
SCRIPT = ROOT / "build_script/pptx_to_h5p_converter pdf v1_4.py"
SOURCE = ROOT / "tests/fixtures/introduction_to_h5p/reference/powerpoint.pdf"
sys.path.insert(0, str(SCRIPT.parent))
spec = importlib.util.spec_from_file_location("pdf_input_cli_v1_4", SCRIPT)
cli = importlib.util.module_from_spec(spec)
spec.loader.exec_module(cli)


def mixed_pages(destination):
    """A landscape red page followed by a portrait blue page, without font dependencies."""
    objects = [
        b"<< /Type /Catalog /Pages 2 0 R >>",
        b"<< /Type /Pages /Kids [3 0 R 5 0 R] /Count 2 >>",
    ]
    for number, width, height, color in ((3, 160, 90, "1 0 0"), (5, 90, 160, "0 0 1")):
        content = f"{color} rg 0 0 {width} {height} re f".encode()
        objects.extend([
            f"<< /Type /Page /Parent 2 0 R /MediaBox [0 0 {width} {height}] "
            f"/Resources << >> /Contents {number + 1} 0 R >>".encode(),
            f"<< /Length {len(content)} >>\nstream\n".encode() + content + b"\nendstream",
        ])
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
    destination.write_bytes(data)


class PdfInputV14Tests(unittest.TestCase):
    def setUp(self):
        self.temporary = tempfile.TemporaryDirectory(prefix="pdf-cli-tests-")
        self.addCleanup(self.temporary.cleanup)
        self.directory = Path(self.temporary.name).resolve()

    def run_cli(self, *arguments, no_site=False):
        return subprocess.run([sys.executable, *(["-S"] if no_site else []), str(SCRIPT),
                               *map(str, arguments)], capture_output=True, text=True, timeout=30)

    def test_help_and_missing_dependencies_without_site_packages(self):
        result = self.run_cli("--help", no_site=True)
        self.assertEqual(result.returncode, 0, result.stderr)
        self.assertIn("--resolution", result.stdout)
        self.assertIn("--app", result.stdout)
        result = self.run_cli(SOURCE, no_site=True)
        self.assertEqual(result.returncode, 1)
        self.assertIn("pip install -r requirements.txt", result.stderr)
        self.assertNotIn("Traceback", result.stderr)

    def test_input_output_validation_and_exit_codes(self):
        not_a_directory = self.directory / "output.h5p"
        not_a_directory.touch()
        for source, output in ((self.directory / "missing.pdf", None),
                               (self.directory, None), (SCRIPT, None),
                               (SOURCE, self.directory), (SOURCE, self.directory / "bad.zip")):
            with self.subTest(source=source, output=output), self.assertRaises(ValueError):
                cli.build_jobs(source, output)
        (self.directory / "document.pdf").touch()
        with self.assertRaisesRegex(ValueError, "must be a directory"):
            cli.build_jobs(self.directory, not_a_directory)
        for arguments in ((SOURCE, "--resolution", "medium"),
                          (SOURCE, "--app", "unknown"),
                          (self.directory / "missing.pdf",)):
            result = self.run_cli(*arguments)
            self.assertEqual(result.returncode, 2, result.stderr)
            self.assertNotIn("Traceback", result.stderr)

    def test_folder_filter_order_and_output_defaults(self):
        for name in ("z.PDF", "A.pdf", "ignored.ppt"):
            (self.directory / name).touch()
        nested = self.directory / "nested.pdf"
        nested.mkdir()
        (nested / "skip.pdf").touch()
        jobs = cli.build_jobs(self.directory)
        self.assertEqual([p.name for p, _ in jobs], ["A.pdf", "z.PDF"])
        self.assertEqual([p.name for _, p in jobs], ["A.h5p", "z.h5p"])
        self.assertTrue(all(p.parent == self.directory for _, p in jobs))
        self.assertEqual(cli.build_jobs(SOURCE), [(SOURCE, SOURCE.with_suffix(".h5p"))])
        output = self.directory / "new directory"
        self.assertTrue(all(p.parent == output for _, p in cli.build_jobs(self.directory, output)))

    def test_colliding_outputs_and_source_symlink_are_rejected(self):
        first, second = self.directory / "same.pdf", self.directory / "same.PDF"
        first.touch()
        second.touch()
        # Simulate a case-sensitive listing on case-insensitive filesystems too.
        with mock.patch.object(Path, "iterdir", return_value=iter((first, second))):
            with self.assertRaisesRegex(ValueError, "same output"):
                cli.build_jobs(self.directory)
        output = self.directory / "source.h5p"
        try:
            output.symlink_to(first)
        except OSError:
            self.skipTest("Symlinks are unavailable")
        with self.assertRaisesRegex(ValueError, "must not replace the source"):
            cli.build_jobs(first, output)

    def test_real_pdf_conversion_preserves_page_order_and_proportions(self):
        from PIL import Image

        source = self.directory / "Mixed pages.pdf"
        mixed_pages(source)
        output = self.directory / "nested/course.h5p"
        for resolution, size in cli.RESOLUTIONS.items():
            with self.subTest(resolution=resolution):
                result = self.run_cli(source, "-o", output, "--resolution", resolution, "--force")
                self.assertEqual(result.returncode, 0, result.stderr)
                self.assertEqual(result.stdout.strip(), str(output))
                self.assertIn("1 succeeded, 0 failed", result.stderr)
                with zipfile.ZipFile(output) as archive:
                    self.assertIsNone(archive.testzip())
                    self.assertEqual(set(archive.namelist()), {
                        "h5p.json", "content/content.json",
                        "content/images/slide_1.png", "content/images/slide_2.png",
                    })
                    metadata = json.loads(archive.read("h5p.json"))
                    self.assertEqual(metadata["title"], source.stem)
                    self.assertEqual(metadata["mainLibrary"], "H5P.CoursePresentation")
                    self.assertIn({"machineName": "H5P.Image", "majorVersion": 1, "minorVersion": 1},
                                  metadata["preloadedDependencies"])
                    slides = json.loads(archive.read("content/content.json"))["presentation"]["slides"]
                    self.assertEqual(len(slides), 2)
                    for index, color in enumerate(((255, 0, 0), (0, 0, 255))):
                        action = slides[index]["elements"][0]["action"]
                        self.assertEqual(action["metadata"]["title"], f"Slide {index + 1}")
                        file = action["params"]["file"]
                        self.assertEqual((file["width"], file["height"]), size)
                        with Image.open(io.BytesIO(archive.read("content/" + file["path"]))) as image:
                            self.assertEqual(image.size, size)
                            self.assertEqual(image.getpixel((size[0] // 2, size[1] // 2)), color)
                            self.assertEqual(image.getpixel((0, size[1] // 2)),
                                             color if index == 0 else (255, 255, 255))

    def test_reference_pdf_and_folder_continuation_without_office(self):
        source = self.directory / "Documents with spaces"
        source.mkdir()
        (source / "A broken.pdf").write_bytes(b"not a PDF")
        shutil.copyfile(SOURCE, source / "B référence.PDF")
        (source / "ignored.ppt").touch()
        output = self.directory / "Converted courses"
        result = self.run_cli(source, "-o", output, "--resolution", "low")
        self.assertEqual(result.returncode, 1, result.stderr)
        self.assertIn("1 succeeded, 1 failed", result.stderr)
        self.assertIn("Failed A broken.pdf", result.stderr)
        self.assertNotIn("Traceback", result.stderr)
        archive_path = output / "B référence.h5p"
        self.assertEqual(result.stdout.strip(), str(archive_path))
        self.assertEqual(list(output.iterdir()), [archive_path])
        with zipfile.ZipFile(archive_path) as archive:
            slides = json.loads(archive.read("content/content.json"))["presentation"]["slides"]
            self.assertEqual(len(slides), 7)

    def test_overwrite_guard_force_and_invalid_pdf_preserve_existing_output(self):
        source = self.directory / "source.pdf"
        mixed_pages(source)
        output = source.with_suffix(".h5p")
        output.write_bytes(b"keep this")
        result = self.run_cli(source)
        self.assertEqual(result.returncode, 1)
        self.assertIn("Use --force", result.stderr)
        self.assertEqual(output.read_bytes(), b"keep this")
        source.write_bytes(b"corrupt PDF")
        result = self.run_cli(source, "--force", "--verbose")
        self.assertEqual(result.returncode, 1)
        self.assertIn("Traceback", result.stderr)
        self.assertEqual(output.read_bytes(), b"keep this")
        mixed_pages(source)
        result = self.run_cli(source, "--force", "--resolution", "low")
        self.assertEqual(result.returncode, 0, result.stderr)
        self.assertTrue(zipfile.is_zipfile(output))
        self.assertEqual(set(self.directory.iterdir()), {source, output})

    def test_packaging_failure_cleans_up_and_preserves_existing_output(self):
        output = self.directory / "course.h5p"
        output.write_bytes(b"original")
        args = cli.create_parser().parse_args([str(SOURCE), "--force", "--resolution", "low"])
        with mock.patch.object(cli, "write_h5p", side_effect=OSError("disk full")) as write:
            with self.assertRaisesRegex(OSError, "disk full"):
                cli.convert(SOURCE, output, args)
        self.assertEqual(output.read_bytes(), b"original")
        self.assertEqual(list(self.directory.iterdir()), [output])
        self.assertFalse(write.call_args.args[2][0][1].parent.exists())

    def test_concurrent_output_creation_is_protected(self):
        output = self.directory / "course.h5p"
        args = cli.create_parser().parse_args([str(SOURCE), "--resolution", "low"])

        def competing_writer(*_):
            output.write_bytes(b"another writer")

        with mock.patch.object(cli, "write_h5p", side_effect=competing_writer):
            with self.assertRaisesRegex(cli.ConversionError, "Output exists"):
                cli.convert(SOURCE, output, args)
        self.assertEqual(output.read_bytes(), b"another writer")
        self.assertEqual(list(self.directory.iterdir()), [output])

    def test_interrupt_cleans_temporary_archive_and_returns_130(self):
        output = self.directory / "course.h5p"
        with mock.patch.object(cli, "write_h5p", side_effect=KeyboardInterrupt) as write, \
                contextlib.redirect_stderr(io.StringIO()) as stderr:
            result = cli.main([str(SOURCE), "-o", str(output), "--resolution", "low"])
        self.assertEqual(result, 130)
        self.assertIn("Interrupted", stderr.getvalue())
        self.assertEqual(list(self.directory.iterdir()), [])
        self.assertFalse(write.call_args.args[2][0][1].parent.exists())


if __name__ == "__main__":
    unittest.main()

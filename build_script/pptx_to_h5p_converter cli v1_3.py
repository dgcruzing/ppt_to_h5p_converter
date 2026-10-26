"""Convert PPTX files to image-based H5P Course Presentations (Python 3.11+)."""

import argparse
import contextlib
import csv
import importlib
import io
import json
import math
import os
from pathlib import Path
import posixpath
import shutil
import signal
import subprocess
import sys
import tempfile
import traceback
import uuid
import xml.etree.ElementTree as ET
import zipfile

_font_verification = importlib.import_module("pptx_to_h5p_converter cli v1_3 verify fonts")
FontCheckError = _font_verification.FontCheckError
missing_fonts = _font_verification.missing_fonts


RESOLUTIONS = {"high": (1920, 1080), "low": (1280, 720)}
PRESENTATION_NS = "http://schemas.openxmlformats.org/presentationml/2006/main"
RELATIONSHIP_NS = "http://schemas.openxmlformats.org/officeDocument/2006/relationships"
DEPENDENCIES = (
    ("H5P.CoursePresentation", 1, 25),
    ("H5P.Image", 1, 1),
    ("FontAwesome", 4, 5),
    ("H5P.FontIcons", 1, 0),
    ("H5P.JoubelUI", 1, 3),
    ("H5P.Transition", 1, 0),
)


class ConversionError(Exception):
    """An actionable conversion failure suitable for display without a traceback."""


def positive_seconds(value):
    try:
        seconds = float(value)
    except ValueError as exc:
        raise argparse.ArgumentTypeError("timeout must be a positive number") from exc
    if not math.isfinite(seconds) or seconds <= 0:
        raise argparse.ArgumentTypeError("timeout must be a finite positive number")
    return seconds


def create_parser():
    parser = argparse.ArgumentParser(
        description="Convert a PPTX file or a folder of PPTX files to H5P.",
        formatter_class=argparse.ArgumentDefaultsHelpFormatter,
    )
    parser.add_argument("input", type=Path, help="PPTX file or folder (not recursive)")
    parser.add_argument("-o", "--output", type=Path,
                        help="output H5P file, or output directory for folder input")
    parser.add_argument("--app", choices=("auto", "powerpoint", "libreoffice"),
                        default="auto", help="app used to render slides")
    parser.add_argument("--libreoffice-path", type=Path,
                        help="LibreOffice executable path if automatic discovery fails")
    parser.add_argument("--include-hidden", action="store_true", help="include hidden slides")
    parser.add_argument("--ignore-missing-fonts", action="store_true",
                        help="skip the font check and its warnings; allow substitute fonts")
    parser.add_argument("--timeout", type=positive_seconds, default=30.0,
                        metavar="SECONDS", help="maximum app export time per presentation")
    parser.add_argument("--resolution", choices=RESOLUTIONS, default="high",
                        help="high: 1920x1080; low: 1280x720")
    parser.add_argument("--force", action="store_true", help="replace existing output files")
    parser.add_argument("--verbose", action="store_true", help="show diagnostics and tracebacks")
    return parser


def select_app(requested, platform=None):
    platform = sys.platform if platform is None else platform
    if requested == "auto":
        return "powerpoint" if platform == "win32" else "libreoffice"
    if requested == "powerpoint" and platform != "win32":
        raise ValueError("PowerPoint automation requires Windows; use --app libreoffice.")
    return requested


def build_jobs(input_path, output_path=None):
    source = input_path.expanduser().resolve()
    output = output_path.expanduser().absolute() if output_path is not None else None
    if source.is_dir():
        sources = sorted(
            (p for p in source.iterdir() if p.is_file() and p.suffix.lower() == ".pptx"
             and not p.name.startswith("~$")),
            key=lambda p: (p.name.casefold(), p.name),
        )
        if not sources:
            raise ValueError(f"No .pptx presentations found in {source} (subfolders are excluded).")
        destination = output if output is not None else source
        if os.path.lexists(destination) and not destination.is_dir():
            raise ValueError("For folder input, --output must be a directory.")
        jobs = [(p, destination / (p.stem + ".h5p")) for p in sources]
    elif source.is_file() and source.suffix.lower() == ".pptx":
        destination = output if output is not None else source.with_suffix(".h5p")
        if destination.suffix.lower() != ".h5p" or destination.is_dir():
            raise ValueError("For single-file input, --output must be an .h5p file path.")
        jobs = [(source, destination)]
    else:
        raise ValueError(f"Input must be an existing .pptx file or folder: {source}")

    seen = set()
    for source, destination in jobs:
        key = str(destination.resolve()).casefold()
        if key in seen:
            raise ValueError(f"Multiple presentations would write to the same output: {destination}")
        if destination.resolve() == source.resolve():
            raise ValueError(f"Output must not replace the source presentation: {source}")
        seen.add(key)
    return jobs


def require_module(name, install_name):
    try:
        return importlib.import_module(name)
    except ImportError as exc:
        raise ConversionError(
            f"Cannot load {install_name}. Activate your virtual environment and run "
            f"python -m pip install -r requirements.txt. Details: {exc}"
        ) from exc


def find_libreoffice(explicit=None):
    if explicit is not None:
        candidate = explicit.expanduser().resolve()
        if candidate.is_file() and os.access(candidate, os.X_OK):
            return candidate
        raise ConversionError(f"Invalid LibreOffice executable: {candidate}")
    for name in ("soffice", "libreoffice"):
        found = shutil.which(name)
        if found:
            # Keep launcher symlinks intact: some launchers use their own location.
            return Path(found).absolute()
    candidates = [Path("/Applications/LibreOffice.app/Contents/MacOS/soffice")]
    for variable in ("ProgramFiles", "ProgramFiles(x86)", "LOCALAPPDATA"):
        if root := os.environ.get(variable):
            base = Path(root) / ("Programs" if variable == "LOCALAPPDATA" else "")
            candidates.extend(base / "LibreOffice" / "program" / name
                              for name in ("soffice.com", "soffice.exe"))
    for candidate in candidates:
        if candidate.is_file() and os.access(candidate, os.X_OK):
            return candidate
    raise ConversionError(
        "LibreOffice was not found. Install LibreOffice, or pass "
        "--libreoffice-path PATH to its soffice executable."
    )


def eligible_slides(source, include_hidden=False):
    """Read actual slide order from presentation relationships, not ZIP filename order."""
    try:
        with zipfile.ZipFile(source) as archive:
            root = ET.fromstring(archive.read("ppt/presentation.xml"))
            relationships = ET.fromstring(archive.read("ppt/_rels/presentation.xml.rels"))
            targets = {
                rel.attrib["Id"]: rel.attrib["Target"] for rel in relationships
                if rel.get("TargetMode") != "External"
            }
            slides = root.find(f"{{{PRESENTATION_NS}}}sldIdLst")
            eligible = []
            for number, slide in enumerate(slides if slides is not None else [], 1):
                target = targets[slide.attrib[f"{{{RELATIONSHIP_NS}}}id"]]
                path = (target.lstrip("/") if target.startswith("/")
                        else posixpath.normpath(posixpath.join("ppt", target)))
                slide_root = ET.fromstring(archive.read(path))
                hidden = slide_root.get("show", "1").lower() in ("0", "false")
                if include_hidden or not hidden:
                    eligible.append(number)
    except (zipfile.BadZipFile, KeyError, ET.ParseError) as exc:
        raise ConversionError(f"Invalid or unreadable PPTX presentation: {source.name}") from exc
    if not eligible:
        raise ConversionError("No eligible slides. Use --include-hidden if the slides are hidden.")
    return eligible


def stop_process(process):
    """Stop only this export's process group/tree, including launcher children."""
    if sys.platform == "win32":
        if process.poll() is None:
            subprocess.run(["taskkill", "/PID", str(process.pid), "/T", "/F"],
                           stdout=subprocess.DEVNULL, stderr=subprocess.DEVNULL, timeout=5)
    else:
        with contextlib.suppress(ProcessLookupError):
            os.killpg(process.pid, signal.SIGTERM)
    try:
        process.communicate(timeout=2)
    except subprocess.TimeoutExpired:
        if sys.platform == "win32":
            process.kill()
        else:
            with contextlib.suppress(ProcessLookupError):
                os.killpg(process.pid, signal.SIGKILL)
        process.communicate(timeout=2)


def run_export(command, timeout, label, verbose=False):
    if verbose:
        print(f"{label} command: {command!r}", file=sys.stderr)
    options = ({"creationflags": subprocess.CREATE_NEW_PROCESS_GROUP}
               if sys.platform == "win32" else {"start_new_session": True})
    try:
        process = subprocess.Popen(command, stdout=subprocess.PIPE, stderr=subprocess.PIPE,
                                   **options)
    except OSError as exc:
        raise ConversionError(f"Cannot start {label}: {exc}") from exc
    try:
        stdout, stderr = process.communicate(timeout=timeout)
    except BaseException as exc:
        stop_process(process)
        if isinstance(exc, subprocess.TimeoutExpired):
            raise ConversionError(
                f"{label} export exceeded {timeout:g} seconds. "
                "Increase --timeout SECONDS for this presentation."
            ) from exc
        raise
    details = (stdout + b"\n" + stderr).decode(errors="replace").strip()
    if verbose and details:
        print(details, file=sys.stderr)
    if process.returncode:
        hint = ""
        if sys.platform == "darwin" and label == "LibreOffice" and process.returncode in (-6, 134):
            hint = (
                " LibreOffice aborted during execution. If running inside a restricted app sandbox, "
                "try the same command in Terminal; macOS app registration can be blocked there."
            )
        raise ConversionError(f"{label} export failed (exit {process.returncode}).{hint} {details[-2000:]}")
    return details


def export_pdf(source, directory, executable, include_hidden, timeout, verbose=False):
    profile = directory / "profile"
    pdf_filter = "pdf:impress_pdf_Export:" + json.dumps({
        "ExportHiddenSlides": {"type": "boolean", "value": include_hidden},
        "ExportNotes": {"type": "boolean", "value": False},
        "ExportNotesPages": {"type": "boolean", "value": False},
        "UseLosslessCompression": {"type": "boolean", "value": True},
        "ReduceImageResolution": {"type": "boolean", "value": False},
    }, separators=(",", ":"))
    details = run_export([
        str(executable), f"-env:UserInstallation={profile.as_uri()}",
        "--headless", "--convert-to", pdf_filter, "--outdir", str(directory), str(source),
    ], timeout, "LibreOffice", verbose)
    pdf = directory / (source.stem + ".pdf")
    # LibreOffice can exit successfully even when a document could not be converted.
    if not pdf.is_file() or pdf.stat().st_size == 0:
        raise ConversionError(f"LibreOffice did not produce a PDF. {details[-1500:]}")
    return pdf


def fit_size(size, canvas):
    factor = min(canvas[0] / size[0], canvas[1] / size[1])
    return tuple(max(1, min(limit, round(value * factor)))
                 for value, limit in zip(size, canvas))


def normalize_image(image, target, destination):
    Image = require_module("PIL.Image", "Pillow")
    # Flatten alpha onto white before converting to RGB.
    with image.convert("RGBA") as rgba:
        with Image.new("RGBA", rgba.size, "white") as background:
            background.alpha_composite(rgba)
            rgb = background.convert("RGB")
    with rgb:
        with rgb.resize(fit_size(rgb.size, target), Image.Resampling.LANCZOS) as fitted:
            with Image.new("RGB", target, "white") as canvas:
                canvas.paste(fitted, ((target[0] - fitted.width) // 2,
                                      (target[1] - fitted.height) // 2))
                canvas.save(destination, "PNG")


def render_pdf(pdf, directory, target, expected_numbers):
    pdfium = require_module("pypdfium2", "pypdfium2")
    images = []
    with pdfium.PdfDocument(pdf) as document:
        if len(document) != len(expected_numbers):
            raise ConversionError(
                f"LibreOffice exported {len(document)} pages; expected {len(expected_numbers)} slides. "
                "Check hidden-slide settings and the source presentation."
            )
        for index, number in enumerate(expected_numbers):
            with contextlib.closing(document[index]) as page:
                scale = min(target[0] / page.get_width(), target[1] / page.get_height())
                # Review comments are PDF annotations, not slide content.
                with contextlib.closing(page.render(scale=scale, draw_annots=False)) as bitmap:
                    with bitmap.to_pil() as image:
                        path = directory / f"slide_{number}.png"
                        normalize_image(image, target, path)
            images.append((number, path))
    return images


def powerpoint_pids():
    """Snapshot existing PowerPoint sessions so they are never terminated by us."""
    result = subprocess.run(
        ["tasklist", "/FI", "IMAGENAME eq POWERPNT.EXE", "/FO", "CSV", "/NH"],
        capture_output=True, check=True, timeout=5,
    )
    return {int(row[1]) for row in csv.reader(io.StringIO(result.stdout.decode(errors="replace")))
            if len(row) > 1 and row[0].casefold() == "powerpnt.exe"}


def export_powerpoint(source, directory, target, include_hidden, ownership_file):
    pythoncom = require_module("pythoncom", "pywin32")
    client = require_module("win32com.client", "pywin32")
    win32process = require_module("win32process", "pywin32")
    existing = powerpoint_pids()
    app = presentation = None
    owned = False
    pythoncom.CoInitialize()
    try:
        try:
            app = client.DispatchEx("PowerPoint.Application")
        except Exception as exc:
            raise ConversionError(
                "Cannot automate PowerPoint. Install and activate desktop Microsoft PowerPoint, "
                f"or use --app libreoffice. Details: {exc}"
            ) from exc
        _, pid = win32process.GetWindowThreadProcessId(app.HWND)
        owned = pid not in existing
        if owned:
            ownership_file.write_text(json.dumps({"pid": pid}), encoding="utf-8")
        presentation = app.Presentations.Open(str(source), ReadOnly=-1, Untitled=-1, WithWindow=0)
        size = (presentation.PageSetup.SlideWidth, presentation.PageSetup.SlideHeight)
        width, height = fit_size(size, target)
        images = []
        for index in range(1, presentation.Slides.Count + 1):
            slide = presentation.Slides.Item(index)
            if not include_hidden and slide.SlideShowTransition.Hidden:
                continue
            path = directory / f"slide_{index}.png"
            slide.Export(str(path), "PNG", width, height)
            images.append((index, path))
        return images
    finally:
        try:
            if presentation is not None:
                presentation.Close()
        finally:
            try:
                if app is not None and owned:
                    app.Quit()
            finally:
                pythoncom.CoUninitialize()


def cleanup_powerpoint(ownership_file):
    if not ownership_file.is_file():
        return
    # A shared, pre-existing PowerPoint PID is never written to this file.
    with contextlib.suppress(OSError, ValueError, subprocess.SubprocessError):
        pid = json.loads(ownership_file.read_text(encoding="utf-8"))["pid"]
        if pid in powerpoint_pids():
            subprocess.run(["taskkill", "/PID", str(pid), "/T", "/F"],
                           stdout=subprocess.DEVNULL, stderr=subprocess.DEVNULL, timeout=5)


def powerpoint_worker(request_path):
    request = json.loads(request_path.read_text(encoding="utf-8"))
    try:
        images = export_powerpoint(
            Path(request["source"]), Path(request["directory"]), tuple(request["target"]),
            request["include_hidden"], Path(request["ownership_file"]),
        )
        Path(request["result_file"]).write_text(json.dumps([n for n, _ in images]), encoding="utf-8")
        return 0
    except Exception as exc:
        print(str(exc), file=sys.stderr)
        if request["verbose"]:
            traceback.print_exc()
        return 1


def render_powerpoint(source, directory, target, expected_numbers, args):
    Image = require_module("PIL.Image", "Pillow")
    ownership = directory / "powerpoint-owner.json"
    result = directory / "powerpoint-result.json"
    request = directory / "powerpoint-request.json"
    request.write_text(json.dumps({
        "source": str(source), "directory": str(directory), "target": target,
        "include_hidden": args.include_hidden, "verbose": args.verbose,
        "ownership_file": str(ownership), "result_file": str(result),
    }), encoding="utf-8")
    try:
        run_export([sys.executable, str(Path(__file__).resolve()), "--_powerpoint-worker", str(request)],
                   args.timeout, "PowerPoint", args.verbose)
    finally:
        cleanup_powerpoint(ownership)
    numbers = json.loads(result.read_text(encoding="utf-8"))
    if numbers != expected_numbers:
        raise ConversionError("PowerPoint's exported slide list does not match the source slide list.")
    images = []
    for number in numbers:
        path = directory / f"slide_{number}.png"
        with Image.open(path) as image:
            # Load before writing to the same filename (especially important on Windows).
            copied = image.copy()
        with copied:
            normalize_image(copied, target, path)
        images.append((number, path))
    return images


def write_h5p(source, archive_path, images, resolution):
    slides = []
    with zipfile.ZipFile(archive_path, "w", zipfile.ZIP_DEFLATED) as archive:
        for position, (number, image) in enumerate(images, 1):
            filename = f"images/slide_{position}.png"
            archive.write(image, f"content/{filename}")
            slides.append({
                "elements": [{
                    "x": 0, "y": 0, "width": 100, "height": 100,
                    "action": {
                        "library": "H5P.Image 1.1",
                        "params": {"file": {
                            "path": filename, "mime": "image/png", "copyright": {"license": "U"},
                            "width": resolution[0], "height": resolution[1],
                        }},
                        "subContentId": str(uuid.uuid4()),
                        "metadata": {"contentType": "Image", "license": "U", "title": f"Slide {number}"},
                    },
                    "alwaysDisplayComments": False, "backgroundOpacity": 0,
                    "displayAsButton": False, "buttonSize": "big", "goToSlideType": "specified",
                    "invisible": False, "solution": "",
                }],
                "slideBackgroundSelector": {},
            })
        content = {"presentation": {
            "slides": slides, "keywordListEnabled": False, "globalBackgroundSelector": {},
            "keywordListAlwaysShow": False, "keywordListAutoHide": False, "keywordListOpacity": 90,
        }}
        metadata = {
            "title": source.stem, "language": "und", "mainLibrary": "H5P.CoursePresentation",
            "embedTypes": ["div"], "license": "U", "defaultLanguage": "en",
            "preloadedDependencies": [
                {"machineName": name, "majorVersion": major, "minorVersion": minor}
                for name, major, minor in DEPENDENCIES
            ],
        }
        archive.writestr("content/content.json", json.dumps(content, ensure_ascii=False))
        archive.writestr("h5p.json", json.dumps(metadata, ensure_ascii=False))


def convert(source, destination, app, executable, args):
    if os.path.lexists(destination) and not args.force:
        raise ConversionError(f"Output exists: {destination}. Use --force to replace it.")
    numbers = eligible_slides(source, args.include_hidden)
    if not args.ignore_missing_fonts:
        try:
            missing = missing_fonts(source, numbers, app, executable)
        except FontCheckError as exc:
            raise ConversionError(f"Font check failed: {exc}\n"
                                  "To continue without font verification, add --ignore-missing-fonts.") from exc
        if missing:
            listing = "\n".join(f"  - {font.family} — {font.style}" for font in missing)
            raise ConversionError(
                f"Missing fonts:\n{listing}\n\n"
                "Install these fonts on this computer before converting. Missing fonts may change "
                "text size, spacing, and line breaks.\n"
                "To continue with substitute fonts, add --ignore-missing-fonts."
            )
    target = RESOLUTIONS[args.resolution]
    destination.parent.mkdir(parents=True, exist_ok=True)
    with tempfile.TemporaryDirectory(prefix="pptx-to-h5p-") as temporary:
        directory = Path(temporary).resolve()
        if app == "libreoffice":
            pdf = export_pdf(source, directory, executable, args.include_hidden,
                             args.timeout, args.verbose)
            images = render_pdf(pdf, directory, target, numbers)
        else:
            images = render_powerpoint(source, directory, target, numbers, args)
        handle, filename = tempfile.mkstemp(prefix=f".{destination.stem}-", suffix=".tmp",
                                           dir=destination.parent)
        os.close(handle)
        temporary_archive = Path(filename)
        try:
            write_h5p(source, temporary_archive, images, target)
            if args.force:
                os.replace(temporary_archive, destination)
            else:
                # Atomic creation also protects a file created by another process during export.
                try:
                    os.link(temporary_archive, destination)
                except FileExistsError as exc:
                    raise ConversionError(f"Output exists: {destination}. Use --force to replace it.") from exc
        finally:
            temporary_archive.unlink(missing_ok=True)


def main(argv=None):
    parser = create_parser()
    args = parser.parse_args(argv)
    if sys.version_info < (3, 11):
        parser.error("Python 3.11 or newer is required.")
    try:
        jobs = build_jobs(args.input, args.output)
        app = select_app(args.app)
        if args.libreoffice_path is not None and app != "libreoffice":
            raise ValueError("--libreoffice-path requires --app libreoffice (or auto on non-Windows).")
    except (ValueError, OSError) as exc:
        parser.error(str(exc))

    try:
        require_module("PIL.Image", "Pillow")
        if app == "libreoffice":
            require_module("pypdfium2", "pypdfium2")
            executable = find_libreoffice(args.libreoffice_path)
        else:
            for module in ("pythoncom", "win32com.client", "win32process"):
                require_module(module, "pywin32")
            executable = None
        print(f"Using {'PowerPoint' if app == 'powerpoint' else 'LibreOffice'}", file=sys.stderr)
        failures = 0
        for source, destination in jobs:
            try:
                convert(source, destination, app, executable, args)
                print(str(destination), flush=True)
            except Exception as exc:
                failures += 1
                print(f"Failed {source.name}: {exc}", file=sys.stderr)
                if args.verbose:
                    traceback.print_exc()
        print(f"{len(jobs) - failures} succeeded, {failures} failed.", file=sys.stderr)
        return 1 if failures else 0
    except KeyboardInterrupt:
        print("Interrupted. Temporary conversion files have been cleaned up.", file=sys.stderr)
        return 130
    except Exception as exc:
        print(f"Error: {exc}", file=sys.stderr)
        if args.verbose:
            traceback.print_exc()
        return 1


if __name__ == "__main__":
    if len(sys.argv) == 3 and sys.argv[1] == "--_powerpoint-worker":
        sys.exit(powerpoint_worker(Path(sys.argv[2])))
    sys.exit(main())

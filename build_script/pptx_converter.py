"""PPTX rendering backend for the v1.4 converter.

The legacy v1.3 entry point stays independent. Only its unchanged font checker
is reused here; PDF-only v1.4 conversion does not import this backend.
"""

import contextlib
import csv
import importlib
import io
import json
import os
from pathlib import Path
import posixpath
import shutil
import signal
import subprocess
import sys
import tempfile
import traceback
import xml.etree.ElementTree as ET
import zipfile

from h5p_converter_common import (
    ConversionError, RESOLUTIONS, atomic_output, fit_size, normalize_image,
    render_pdf, require_module, write_h5p,
)

_font_verification = importlib.import_module("pptx_to_h5p_converter cli v1_3 verify fonts")
FontCheckError = _font_verification.FontCheckError
missing_fonts = _font_verification.missing_fonts

PRESENTATION_NS = "http://schemas.openxmlformats.org/presentationml/2006/main"
RELATIONSHIP_NS = "http://schemas.openxmlformats.org/officeDocument/2006/relationships"


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
        with atomic_output(destination, args.force) as archive:
            write_h5p(source, archive, images, target)

if __name__ == "__main__":
    if len(sys.argv) == 3 and sys.argv[1] == "--_powerpoint-worker":
        sys.exit(powerpoint_worker(Path(sys.argv[2])))
    sys.exit("This module is an internal PPTX backend. Run the v1.4 converter CLI.")

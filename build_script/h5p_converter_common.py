"""Shared PDF rendering, image normalization, and H5P packaging helpers.

Third-party dependencies are loaded only when needed, so CLI help works without
site packages and PDF conversion does not load any Office or font-check code.
"""

import argparse
import contextlib
import importlib
import json
import math
import os
from pathlib import Path
import sys
import tempfile
import uuid
import zipfile

RESOLUTIONS = {"high": (1920, 1080), "low": (1280, 720)}


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


def build_jobs(input_path, output_path=None, extensions=(".pptx", ".pdf")):
    """Plan a non-recursive batch and reject conflicting outputs before rendering."""
    source = input_path.expanduser().resolve()
    output = output_path.expanduser().absolute() if output_path is not None else None
    formats = " or ".join(extensions)

    def supported(path):
        return (path.is_file() and path.suffix.lower() in extensions
                and not (path.suffix.lower() == ".pptx" and path.name.startswith("~$")))

    if source.is_dir():
        sources = sorted((p for p in source.iterdir() if supported(p)),
                         key=lambda p: (p.name.casefold(), p.name))
        if not sources:
            raise ValueError(f"No {formats} documents found in {source} (subfolders are excluded).")
        destination = output if output is not None else source
        if os.path.lexists(destination) and not destination.is_dir():
            raise ValueError("For folder input, --output must be a directory.")
        jobs = [(p, destination / (p.stem + ".h5p")) for p in sources]
    elif supported(source):
        destination = output if output is not None else source.with_suffix(".h5p")
        if destination.suffix.lower() != ".h5p" or destination.is_dir():
            raise ValueError("For single-file input, --output must be an .h5p file path.")
        jobs = [(source, destination)]
    else:
        raise ValueError(f"Input must be an existing {formats} file or folder: {source}")

    sources = {str(path.resolve()).casefold() for path, _ in jobs}
    seen = set()
    for _, destination in jobs:
        key = str(destination.resolve()).casefold()
        if key in seen:
            raise ValueError(f"Multiple documents would write to the same output: {destination}")
        if key in sources:
            raise ValueError(f"Output must not replace the source document: {destination}")
        seen.add(key)
    return jobs


@contextlib.contextmanager
def atomic_output(destination, force=False):
    """Publish only a complete archive, preserving existing output on failures."""
    if os.path.lexists(destination) and not force:
        raise ConversionError(f"Output exists: {destination}. Use --force to replace it.")
    destination.parent.mkdir(parents=True, exist_ok=True)
    handle, filename = tempfile.mkstemp(prefix=f".{destination.stem}-", suffix=".tmp",
                                       dir=destination.parent)
    os.close(handle)
    temporary_archive = Path(filename)
    try:
        yield temporary_archive
        if force:
            os.replace(temporary_archive, destination)
        else:
            # The destination may have been created by another process during rendering.
            try:
                os.link(temporary_archive, destination)
            except FileExistsError as exc:
                raise ConversionError(f"Output exists: {destination}. Use --force to replace it.") from exc
    finally:
        temporary_archive.unlink(missing_ok=True)


def positive_seconds(value):
    try:
        seconds = float(value)
    except ValueError as exc:
        raise argparse.ArgumentTypeError("timeout must be a positive number") from exc
    if not math.isfinite(seconds) or seconds <= 0:
        raise argparse.ArgumentTypeError("timeout must be a finite positive number")
    return seconds


def require_module(name, install_name):
    try:
        return importlib.import_module(name)
    except ImportError as exc:
        raise ConversionError(
            f"Cannot load {install_name}. Activate your virtual environment and run "
            f"python -m pip install -r requirements.txt. Details: {exc}"
        ) from exc


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


def render_pdf(pdf, directory, target, expected_numbers=None):
    """Rasterize a PDF; optionally verify slide numbers from a PPTX export."""
    pdfium = require_module("pypdfium2", "pypdfium2")
    images = []
    with pdfium.PdfDocument(pdf) as document:
        if expected_numbers is None:
            expected_numbers = range(1, len(document) + 1)
            if not expected_numbers:
                raise ConversionError("The PDF contains no pages.")
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


def select_app(requested, platform=None):
    platform = sys.platform if platform is None else platform
    if requested == "auto":
        return "powerpoint" if platform == "win32" else "libreoffice"
    if requested == "powerpoint" and platform != "win32":
        raise ValueError("PowerPoint automation requires Windows; use --app libreoffice.")
    return requested

"""Convert PPTX, PDF, or mixed folders to image-based H5P (Python 3.11+)."""

import argparse
import importlib
import os
from pathlib import Path
import sys
import tempfile
import traceback

from h5p_converter_common import (
    ConversionError, RESOLUTIONS, atomic_output, build_jobs, positive_seconds,
    render_pdf, require_module, select_app, write_h5p,
)


def create_parser():
    parser = argparse.ArgumentParser(
        description="Convert a PPTX or PDF file, or a folder containing both, to H5P.",
        epilog=("Presentation options apply only to PPTX files and are ignored for PDFs. "
                "PDFs require no Office installation. Legacy .ppt files are not supported."),
        formatter_class=argparse.ArgumentDefaultsHelpFormatter,
    )
    parser.add_argument("input", type=Path, help="PPTX/PDF file or mixed folder (not recursive)")
    parser.add_argument("-o", "--output", type=Path,
                        help="output H5P file, or output directory for folder input")
    parser.add_argument("--resolution", choices=RESOLUTIONS, default="high",
                        help="high: 1920x1080; low: 1280x720")
    parser.add_argument("--force", action="store_true", help="replace existing output files")
    parser.add_argument("--verbose", action="store_true", help="show diagnostics and tracebacks")
    presentation = parser.add_argument_group("PPTX presentation options")
    presentation.add_argument("--app", choices=("auto", "powerpoint", "libreoffice"),
                              default="auto", help="app used to render PPTX slides")
    presentation.add_argument("--libreoffice-path", type=Path,
                              help="LibreOffice executable path if automatic discovery fails")
    presentation.add_argument("--include-hidden", action="store_true", help="include hidden PPTX slides")
    presentation.add_argument("--ignore-missing-fonts", action="store_true",
                              help="skip PPTX font verification and allow substitute fonts")
    presentation.add_argument("--timeout", type=positive_seconds, default=30.0,
                              metavar="SECONDS", help="maximum app export time per PPTX presentation")
    return parser


def prepare_pptx(app, args):
    """Load the presentation backend only when a PPTX job needs it."""
    engine = importlib.import_module("pptx_converter")
    if app == "libreoffice":
        require_module("pypdfium2", "pypdfium2")
        executable = engine.find_libreoffice(args.libreoffice_path)
    else:
        for module in ("pythoncom", "win32com.client", "win32process"):
            require_module(module, "pywin32")
        executable = None
    print(f"Using {'PowerPoint' if app == 'powerpoint' else 'LibreOffice'} for PPTX files", file=sys.stderr)
    return engine, executable


def convert(source, destination, args, app=None, backend=None):
    if os.path.lexists(destination) and not args.force:
        raise ConversionError(f"Output exists: {destination}. Use --force to replace it.")
    if source.suffix.lower() == ".pdf":
        target = RESOLUTIONS[args.resolution]
        if args.verbose:
            print(f"Rendering {source.name} with PDFium", file=sys.stderr)
        with tempfile.TemporaryDirectory(prefix="pdf-to-h5p-") as temporary:
            images = render_pdf(source, Path(temporary).resolve(), target)
            with atomic_output(destination, args.force) as archive:
                write_h5p(source, archive, images, target)
    elif source.suffix.lower() == ".pptx":
        app = select_app(args.app) if app is None else app
        engine, executable = prepare_pptx(app, args) if backend is None else backend
        engine.convert(source, destination, app, executable, args)
    else:
        raise ConversionError(f"Unsupported input format: {source.suffix}. Use .pptx or .pdf.")


def main(argv=None):
    parser = create_parser()
    args = parser.parse_args(argv)
    if sys.version_info < (3, 11):
        parser.error("Python 3.11 or newer is required.")
    try:
        jobs = build_jobs(args.input, args.output)
        app = None
        if any(source.suffix.lower() == ".pptx" for source, _ in jobs):
            app = select_app(args.app)
            if args.libreoffice_path is not None and app != "libreoffice":
                raise ValueError("--libreoffice-path requires --app libreoffice (or auto on non-Windows).")
    except (ValueError, OSError) as exc:
        parser.error(str(exc))

    try:
        require_module("PIL.Image", "Pillow")
        backend = None
        backend_error = None
        failures = 0
        for source, destination in jobs:
            try:
                if source.suffix.lower() == ".pptx":
                    # Check before setup, preserving the existing-output diagnostic.
                    if os.path.lexists(destination) and not args.force:
                        raise ConversionError(f"Output exists: {destination}. Use --force to replace it.")
                    if backend_error is not None:
                        raise ConversionError(str(backend_error)) from backend_error
                    if backend is None:
                        try:
                            backend = prepare_pptx(app, args)
                        except Exception as exc:
                            backend_error = exc
                            raise
                convert(source, destination, args, app, backend)
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
    sys.exit(main())

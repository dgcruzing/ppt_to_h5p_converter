import hashlib
import logging
import os
import tempfile
import traceback
import pythoncom
import win32com.client


def generate_slide_hash(ppt_path: str, slide_index: int) -> str:
    """Generates a 12-character unique hex hash for a slide based on source file

    path, file modification time, and 1-based slide index.
    """
    abs_path = os.path.abspath(ppt_path)
    
    # Get last modified time of the file to invalidate cache if PPTX changes
    mtime = os.path.getmtime(abs_path) if os.path.exists(abs_path) else 0.0

    # Compose unique string seed
    seed = f"{abs_path}|{mtime}|{slide_index}"
    
    # Return first 12 characters of SHA-256 digest
    return hashlib.sha256(seed.encode("utf-8")).hexdigest()[:12]


def export_slides_to_images(
    ppt_path: str, resolution: tuple = (1920, 1080)
) -> tuple[str, list[str]]:
    """Opens a PowerPoint file silently via COM automation and exports every slide

    as a PNG image into a temporary directory using deterministic unique hashes.
    """
    if not os.path.isabs(ppt_path):
        ppt_path = os.path.abspath(ppt_path)

    if not os.path.exists(ppt_path):
        raise FileNotFoundError(f"PowerPoint file not found: {ppt_path}")

    temp_dir = tempfile.mkdtemp(prefix="ppt_h5p_")
    image_paths = []

    pythoncom.CoInitialize()

    powerpoint = None
    presentation = None

    try:
        powerpoint = win32com.client.Dispatch("PowerPoint.Application")
        presentation = powerpoint.Presentations.Open(
            ppt_path, ReadOnly=True, Untitled=False, WithWindow=False
        )

        logging.info(f"Opened presentation: {ppt_path}")

        for i, slide in enumerate(presentation.Slides):
            slide_idx = i + 1
            # Generate unique hash for this specific slide
            slide_hash = generate_slide_hash(ppt_path, slide_idx)
            
            # Filename format: slide_001_a1b2c3d4e5f6.png
            image_filename = f"slide_{slide_idx:03d}_{slide_hash}.png"
            image_path = os.path.join(temp_dir, image_filename)

            logging.debug(
                f"Exporting slide {slide_idx}/{presentation.Slides.Count} to {image_filename}"
            )

            slide.Export(image_path, "PNG", resolution[0], resolution[1])
            image_paths.append(image_path)

        return temp_dir, image_paths

    except Exception as e:
        logging.error(f"Failed to export PowerPoint slides: {str(e)}")
        logging.debug(traceback.format_exc())
        raise RuntimeError(f"PowerPoint slide export failed: {str(e)}") from e

    finally:
        if presentation:
            try:
                presentation.Close()
            except Exception:
                pass
        if powerpoint:
            try:
                powerpoint.Quit()
            except Exception:
                pass
        pythoncom.CoUninitialize()
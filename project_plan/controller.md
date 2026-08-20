Now that `gui.py` is established as the View layer, execution flows cleanly between three primary modules across the project lifecycle: **`controller.py`**, **`converter/ppt_exporter.py`**, and **`converter/h5p_builder.py`**.

```
                           +----------------------+
                           |      controller.py   |
                           +----------+-----------+
                                      |
             +------------------------+------------------------+
             |                                                 |
             v                                                 v
  +--------------------+                             +-------------------+
  |       gui.py       |                             |    converter/     |
  |    (View Layer)    |                             |  ppt_exporter.py  |
  +--------------------+                             +---------+---------+
                                                               |
                                                               v
                                                     +-------------------+
                                                     |    converter/     |
                                                     |   h5p_builder.py  |
                                                     +-------------------+

```

---

### 1. Methods Executed on `gui.py` (Called by `controller.py`)

The controller listens to GUI triggers and updates the UI state before, during, and after processing.

| Method                                  | Caller          | Purpose                                                                                                 |
| --------------------------------------- | --------------- | ------------------------------------------------------------------------------------------------------- |
| `get_queue_data()`                      | `controller.py` | Reads all queued PowerPoint files, resolution choices, and output paths when starting conversion.       |
| `set_processing_state(True / False)`    | `controller.py` | Disables buttons (Add, Clear, Convert) during running tasks and re-enables them on completion.          |
| `update_item_status(item_id, status)`   | `controller.py` | Updates row statuses in real-time (e.g., `Pending` → `Processing...` → `Done` / `Failed`).              |
| `update_progress(current, total, text)` | `controller.py` | Advances the global progress bar and updates text (e.g., `"Processing file 2 of 5..."`).                |
| `append_log(message)`                   | `controller.py` | Streams informational messages, timing data, or exception stack traces into the collapsible log drawer. |

---

### 2. Methods Called in `converter/ppt_exporter.py`

This module manages Windows PowerPoint COM interaction (`win32com` / `pythoncom`) and converts `.pptx` slides into intermediate high-resolution PNG images.

| Method                            | Called By                   | Purpose                                                                             |
| --------------------------------- | --------------------------- | ----------------------------------------------------------------------------------- |
| `pythoncom.CoInitialize()`        | `export_slides_to_images()` | Initializes the COM engine for the current background execution thread.             |
| `powerpoint.Presentations.Open()` | `export_slides_to_images()` | Opens the specified PowerPoint presentation silently in the background.             |
| `slide.Export(path, "PNG", w, h)` | `export_slides_to_images()` | Loops over slides and renders each as a `.png` file into a temporary directory.     |
| `pythoncom.CoUninitialize()`      | `export_slides_to_images()` | Safely uninitializes the COM library in a `finally` block to prevent process leaks. |

**Exported Public Function:**

- `export_slides_to_images(ppt_path: str, resolution: tuple) -> (temp_dir, list_of_image_paths)`

---

### 3. Methods Called in `converter/h5p_builder.py`

This module accepts exported image paths and constructs the valid H5P archive (`.h5p` ZIP containing `h5p.json`, `content/content.json`, and slide media).

| Method / Call                                    | Called By             | Purpose                                                                  |
| ------------------------------------------------ | --------------------- | ------------------------------------------------------------------------ |
| `Image.open(img_path).size`                      | `build_h5p_package()` | Reads pixel dimensions for each PNG slide using Pillow.                  |
| `zipfile.ZipFile(output_path, 'w')`              | `build_h5p_package()` | Creates the `.h5p` compressed ZIP package.                               |
| `zip_file.writestr('content/content.json', ...)` | `build_h5p_package()` | Writes the H5P Course Presentation metadata and slide layout JSON specs. |
| `zip_file.writestr('h5p.json', ...)`             | `build_h5p_package()` | Writes the main library package dependency declaration.                  |
| `shutil.rmtree(temp_dir)`                        | `build_h5p_package()` | Cleans up the temporary directory containing intermediate `.png` files.  |

**Exported Public Function:**

- `build_h5p_package(image_paths: list, output_h5p_path: str, title: str) -> None`
- `convert_ppt_to_h5p(ppt_path: str, output_path: str, resolution: tuple)` _(High-level wrapper chaining `ppt_exporter` and `h5p_builder`)_

---

### 4. Controller Method Execution Flow (`controller.py`)

`controller.py` glues everything together on a background worker thread:

```python
# Execution Sequence Inside Background Thread

1. items = view.get_queue_data()
2. view.set_processing_state(True)

3. for idx, item in enumerate(items):
       view.update_item_status(item["item_id"], "Exporting Slides...")
       view.append_log(f"Opening {item['source_path']}...")

       # Step A: Export slides via PowerPoint COM
       temp_dir, image_paths = ppt_exporter.export_slides_to_images(
           item["source_path"], item["resolution"]
       )

       # Step B: Package images into H5P zip structure
       view.update_item_status(item["item_id"], "Building H5P...")
       h5p_builder.build_h5p_package(image_paths, item["output_path"])

       # Step C: Update Progress
       view.update_item_status(item["item_id"], "Completed")
       view.update_progress(idx + 1, len(items))

4. view.set_processing_state(False)

```

The best next step is **`converter/ppt_exporter.py`**.

Building the converter modules before the controller allows you to test the core PowerPoint conversion and H5P creation logic independently before wiring everything up to the GUI.

Here is the natural order to finish the modular architecture:

1. **`converter/ppt_exporter.py`** _(Recommended Next)_
   Extracts PowerPoint slides to temporary image files via COM automation.
2. **`converter/h5p_builder.py`**
   Takes those images, constructs the H5P JSON metadata, and packages them into a `.h5p` archive.
3. **`controller.py`**
   Binds `gui.py` to the conversion modules using background worker threads, driving real-time queue and progress bar updates.
4. **`main.py`**
   The lightweight app entry point that instantiates and starts the controller.

---

### Implementation: `converter/ppt_exporter.py`

Create `converter/ppt_exporter.py` with the following code:

```python
import logging
import os
import tempfile
import traceback
import pythoncom
import win32com.client


def export_slides_to_images(
    ppt_path: str, resolution: tuple = (1920, 1080)
) -> tuple[str, list[str]]:
    """Opens a PowerPoint file silently via COM automation and exports every slide

    as a PNG image into a temporary directory.

    Args:
        ppt_path: Absolute path to the source .pptx file.
        resolution: Target (width, height) tuple in pixels.

    Returns:
        A tuple containing:
            - temp_dir (str): Path to the temporary folder holding images.
            - image_paths (list[str]): Sorted list of absolute paths to generated PNGs.
    """
    if not os.path.isabs(ppt_path):
        ppt_path = os.path.abspath(ppt_path)

    if not os.path.exists(ppt_path):
        raise FileNotFoundError(f"PowerPoint file not found: {ppt_path}")

    temp_dir = tempfile.mkdtemp(prefix="ppt_h5p_")
    image_paths = []

    # Initialize COM library for the current thread
    pythoncom.CoInitialize()

    powerpoint = None
    presentation = None

    try:
        powerpoint = win32com.client.Dispatch("PowerPoint.Application")
        presentation = powerpoint.Presentations.Open(
            ppt_path, ReadOnly=True, Untitled=False, WithWindow=False
        )

        logging.info(f"Opened presentation: {ppt_path}")
        logging.info(f"Total slides: {presentation.Slides.Count}")

        for i, slide in enumerate(presentation.Slides):
            image_path = os.path.join(temp_dir, f"slide_{i+1:03d}.png")
            logging.debug(
                f"Exporting slide {i+1}/{presentation.Slides.Count} to {image_path}"
            )

            # Export slide as PNG
            slide.Export(image_path, "PNG", resolution[0], resolution[1])
            image_paths.append(image_path)

        return temp_dir, image_paths

    except Exception as e:
        logging.error(f"Failed to export PowerPoint slides: {str(e)}")
        logging.debug(traceback.format_exc())
        raise RuntimeError(
            f"PowerPoint slide export failed: {str(e)}"
        ) from e

    finally:
        # Clean up PowerPoint COM objects safely
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

```

import os
import shutil
from converter.ppt_exporter import export_slides_to_images

# Path to a real PowerPoint file on your machine
SAMPLE_PPTX = r"C:\Workspace\pptx2h5p\pptx2h5p\test_ppt\1-3_Windows Server Overview.pptx"

if __name__ == "__main__":
    if not os.path.exists(SAMPLE_PPTX):
        print(f"Error: Please update SAMPLE_PPTX with a valid path. '{SAMPLE_PPTX}' not found.")
        exit(1)

    print("Starting PowerPoint slide export...")
    try:
        temp_dir, image_paths = export_slides_to_images(SAMPLE_PPTX, resolution=(1920, 1080))
        
        print(f"\nSuccess!")
        print(f"Temporary Directory: {temp_dir}")
        print(f"Total Slides Exported: {len(image_paths)}")
        print("Generated files:")
        for path in image_paths:
            print(f" - {path} (Exists: {os.path.exists(path)})")

    except Exception as e:
        print(f"\nExport Failed: {e}")

    finally:
        # Cleanup temporary files after inspection
        if 'temp_dir' in locals() and os.path.exists(temp_dir):
            shutil.rmtree(temp_dir)
            print("\nCleaned up temporary directory.")
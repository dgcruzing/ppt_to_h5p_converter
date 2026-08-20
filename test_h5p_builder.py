import json
import os
import shutil
import tempfile
import zipfile
from PIL import Image
from converter.h5p_builder import build_h5p_package


def run_manual_test():
    # 1. Create a temporary directory and 2 synthetic slide images
    temp_dir = tempfile.mkdtemp(prefix="h5p_test_")
    img1_path = os.path.join(temp_dir, "slide_001_a1b2c3d4e5f6.png")
    img2_path = os.path.join(temp_dir, "slide_002_f6e5d4c3b2a1.png")

    Image.new("RGB", (1920, 1080), color="blue").save(img1_path)
    Image.new("RGB", (1920, 1080), color="red").save(img2_path)

    output_h5p = os.path.join(temp_dir, "output_presentation.h5p")

    try:
        print("Building H5P package...")
        build_h5p_package(
            image_paths=[img1_path, img2_path],
            output_h5p_path=output_h5p,
            title="Test Presentation",
        )

        assert os.path.exists(output_h5p), "H5P file was not created!"
        print(
            f"Package created successfully: {output_h5p} ({os.path.getsize(output_h5p)} bytes)"
        )

        # 2. Inspect ZIP contents to ensure valid schema
        with zipfile.ZipFile(output_h5p, "r") as z:
            file_list = z.namelist()
            print("\nArchive contents:")
            for f in file_list:
                print(f" - {f}")

            # Assert key required files exist
            assert "h5p.json" in file_list
            assert "content/content.json" in file_list
            assert "content/images/slide_001_a1b2c3d4e5f6.png" in file_list
            assert "content/images/slide_002_f6e5d4c3b2a1.png" in file_list

            # Verify content.json is valid JSON
            content_json = json.loads(z.read("content/content.json"))
            slides_count = len(content_json["presentation"]["slides"])
            assert slides_count == 2, (
                f"Expected 2 slides in JSON, found {slides_count}"
            )

        print(
            "\nTest passed successfully! Archive structure and JSON data are valid."
        )

    finally:
        # Cleanup test files
        shutil.rmtree(temp_dir)


if __name__ == "__main__":
    run_manual_test()
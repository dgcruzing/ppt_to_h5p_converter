# ppt_to_h5p_converter
PowerPoint to H5P Converter
This tool converts PowerPoint presentations (.pptx files) to H5P course presentations. It creates high-resolution PNG images of each slide and packages them into an H5P file format.
System Requirements

Windows operating system
Microsoft PowerPoint installed

Installation

Go to the Releases page of this repository.
Download the latest ppt_to_h5p_converter.exe file.
Save the file to a location on your computer where you have write permissions.

Usage

Double-click the ppt_to_h5p_converter.exe file to run the application.
Use the GUI to select your input PowerPoint file and specify the output H5P file.
Click "Convert" to start the conversion process.
Check the output directory for your converted H5P file.

Troubleshooting
If you encounter any issues:

Check that you have Microsoft PowerPoint installed and up to date.
Ensure you have write permissions in the directory where you're trying to save the H5P file.
Look for a ppt_to_h5p_converter.log file in the same directory as the executable. This log file may contain detailed error messages.

If problems persist, please open an issue on this repository with a description of the problem and the contents of the log file.
Contributing
Contributions, issues, and feature requests are welcome. Feel free to check the issues page if you want to contribute.
License
This project is licensed under the MIT License - see the LICENSE file for details.

Look to issues as I am getting false positives, plus they are converting a bit big at the moment. 

# PowerPoint to H5P Converter v1.2 Changelog

## New Features

1. **Resolution Options**: 
   - Added the ability to choose between two resolution options for slide conversion:
     - High Resolution: 1920x1080
     - Low Resolution: 1280x720
   - Implemented radio buttons in the GUI for easy selection of resolution.

2. **GUI Enhancements**:
   - Updated the graphical user interface to include resolution selection options.
   - Improved layout for better user experience.

## Code Changes

1. **convert_slides_to_images Function**:
   - Modified to accept a `resolution` parameter.
   - Now uses the specified resolution when exporting slides to images.

2. **convert_ppt_to_h5p Function**:
   - Updated to pass the selected resolution to `convert_slides_to_images`.

3. **create_gui Function**:
   - Added radio buttons for resolution selection.
   - Modified the `convert` function to use the selected resolution.

## Executable Creation

- Implemented the ability to create a standalone executable (.exe) file using PyInstaller.
- Added instructions for creating the executable, including necessary PyInstaller commands.

## Other Improvements

- Updated error handling and logging to account for resolution-related issues.
- Refined code comments for better maintainability.
- Ensured compatibility with the new resolution options throughout the script.

## Notes

- The executable version requires Microsoft PowerPoint to be installed on the user's system.
- Users should be aware that high-resolution conversions may take longer and produce larger file sizes.

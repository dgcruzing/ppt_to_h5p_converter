import tkinter as tk
from tkinter import filedialog, messagebox
import win32com.client
import json
import os
import zipfile
import uuid
import tempfile
import sys
import logging
import traceback
from PIL import Image
import pythoncom

# Logging configuration
logging.basicConfig(filename='ppt_to_h5p_converter.log', level=logging.DEBUG, 
                    format='%(asctime)s - %(levelname)s - %(message)s')

def resource_path(relative_path):
    try:
        base_path = sys._MEIPASS
    except Exception:
        base_path = os.path.abspath(".")
    return os.path.join(base_path, relative_path)

def convert_slides_to_images(ppt_file):
    temp_dir = tempfile.mkdtemp()
    image_paths = []

    try:
        pythoncom.CoInitialize()
        powerpoint = win32com.client.Dispatch("Powerpoint.Application")
        presentation = powerpoint.Presentations.Open(os.path.abspath(ppt_file), ReadOnly=True)
        
        logging.info(f"Successfully opened PowerPoint file: {ppt_file}")
        logging.info(f"Number of slides: {presentation.Slides.Count}")
        
        for i, slide in enumerate(presentation.Slides):
            image_path = os.path.join(temp_dir, f"slide_{i+1}.png")
            logging.debug(f"Attempting to export slide {i+1} to {image_path}")
            slide.Export(image_path, "PNG", 3840, 2160)  # Export as high-res PNG (4K resolution)
            image_paths.append(image_path)
            logging.debug(f"Successfully exported slide {i+1}")
        
        presentation.Close()
        powerpoint.Quit()
        pythoncom.CoUninitialize()
    except Exception as e:
        logging.error(f"Error in PowerPoint conversion: {str(e)}")
        logging.debug(traceback.format_exc())
        try:
            presentation.Close()
            powerpoint.Quit()
        except:
            pass
        pythoncom.CoUninitialize()
        raise

    return temp_dir, image_paths

def convert_ppt_to_h5p(ppt_file, h5p_file):
    try:
        temp_dir, image_paths = convert_slides_to_images(ppt_file)
    except Exception as e:
        logging.error(f"Failed to convert slides to images: {str(e)}")
        raise

    content = {
        "presentation": {
            "slides": [],
            "keywordListEnabled": False,
            "globalBackgroundSelector": {},
            "keywordListAlwaysShow": False,
            "keywordListAutoHide": False,
            "keywordListOpacity": 90
        }
    }
    
    try:
        with zipfile.ZipFile(h5p_file, 'w', zipfile.ZIP_DEFLATED) as h5p_zip:
            for i, image_path in enumerate(image_paths):
                with Image.open(image_path) as img:
                    width, height = img.size
                
                slide_filename = f"images/slide_{i+1}.png"
                h5p_zip.write(image_path, f"content/{slide_filename}")
                
                h5p_slide = {
                    "elements": [{
                        "x": 0, "y": 0, "width": 100, "height": 100,
                        "action": {
                            "library": "H5P.Image 1.1",
                            "params": {
                                "file": {
                                    "path": slide_filename,
                                    "mime": "image/png",
                                    "copyright": {"license": "U"},
                                    "width": width, "height": height
                                }
                            },
                            "subContentId": str(uuid.uuid4()),
                            "metadata": {"contentType": "Image", "license": "U", "title": f"Slide {i+1}"}
                        },
                        "alwaysDisplayComments": False,
                        "backgroundOpacity": 0,
                        "displayAsButton": False,
                        "buttonSize": "big",
                        "goToSlideType": "specified",
                        "invisible": False,
                        "solution": ""
                    }],
                    "slideBackgroundSelector": {}
                }
                content["presentation"]["slides"].append(h5p_slide)
            
            h5p_zip.writestr('content/content.json', json.dumps(content))
            
            h5p_metadata = {
                "title": os.path.splitext(os.path.basename(ppt_file))[0],
                "language": "und",
                "mainLibrary": "H5P.CoursePresentation",
                "embedTypes": ["div"],
                "license": "U",
                "defaultLanguage": "en",
                "preloadedDependencies": [
                    {"machineName": "H5P.CoursePresentation", "majorVersion": "1", "minorVersion": "25"},
                    {"machineName": "FontAwesome", "majorVersion": "4", "minorVersion": "5"},
                    {"machineName": "H5P.FontIcons", "majorVersion": "1", "minorVersion": "0"},
                    {"machineName": "H5P.JoubelUI", "majorVersion": "1", "minorVersion": "3"},
                    {"machineName": "H5P.Transition", "majorVersion": "1", "minorVersion": "0"}
                ]
            }
            h5p_zip.writestr('h5p.json', json.dumps(h5p_metadata))
    except Exception as e:
        logging.error(f"Error creating H5P file: {str(e)}")
        logging.debug(traceback.format_exc())
        raise
    finally:
        for image_path in image_paths:
            try:
                os.remove(image_path)
            except Exception as e:
                logging.warning(f"Failed to remove temporary file {image_path}: {str(e)}")
        try:
            os.rmdir(temp_dir)
        except Exception as e:
            logging.warning(f"Failed to remove temporary directory {temp_dir}: {str(e)}")
    
    logging.info(f"Conversion complete. Output saved to {h5p_file}")

def create_gui():
    def select_ppt_file():
        file_path = filedialog.askopenfilename(filetypes=[("PowerPoint files", "*.pptx")])
        if file_path:
            input_entry.delete(0, tk.END)
            input_entry.insert(0, file_path)

    def select_output_file():
        file_path = filedialog.asksaveasfilename(defaultextension=".h5p", filetypes=[("H5P files", "*.h5p")])
        if file_path:
            output_entry.delete(0, tk.END)
            output_entry.insert(0, file_path)

    def convert():
        input_file = input_entry.get()
        output_file = output_entry.get()
        
        if not input_file or not output_file:
            messagebox.showerror("Error", "Please select both input and output files.")
            return
        
        if not os.path.exists(input_file):
            messagebox.showerror("Error", f"Input file '{input_file}' does not exist.")
            return
        
        try:
            convert_ppt_to_h5p(input_file, output_file)
            messagebox.showinfo("Success", f"Conversion complete. Output saved to {output_file}")
        except Exception as e:
            error_message = f"An error occurred during conversion: {str(e)}\n\nSee log file for details."
            logging.error(error_message)
            logging.debug(traceback.format_exc())
            messagebox.showerror("Error", error_message)

    root = tk.Tk()
    root.title("PowerPoint to H5P Converter")

    tk.Label(root, text="Select PowerPoint file:").grid(row=0, column=0, sticky="w", padx=5, pady=5)
    input_entry = tk.Entry(root, width=50)
    input_entry.grid(row=0, column=1, padx=5, pady=5)
    tk.Button(root, text="Browse", command=select_ppt_file).grid(row=0, column=2, padx=5, pady=5)

    tk.Label(root, text="Select output H5P file:").grid(row=1, column=0, sticky="w", padx=5, pady=5)
    output_entry = tk.Entry(root, width=50)
    output_entry.grid(row=1, column=1, padx=5, pady=5)
    tk.Button(root, text="Browse", command=select_output_file).grid(row=1, column=2, padx=5, pady=5)

    tk.Button(root, text="Convert", command=convert).grid(row=2, column=1, pady=20)

    root.mainloop()

if __name__ == "__main__":
    create_gui()

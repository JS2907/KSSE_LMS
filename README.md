# KSSE PPT Automation

This repository includes a simple tool to convert an image into a PowerPoint slide. The script analyzes the image for basic shapes, text and even simple tables, then recreates them on a slide using `python-pptx`.

## Requirements

All dependencies are listed in `requirements.txt`. Make sure you have Tesseract installed on your system for OCR.

## Usage

```bash
python backend/image_to_ppt.py path/to/image.png -o output.pptx
```

The generated `output.pptx` will contain a single slide with detected shapes, text and tables positioned to roughly match the source image.

## Windows GUI

For an easier workflow on Windows, launch the simple Tkinter interface:

```bash
python frontend/gui.py
```

Use the *Browse* buttons to select an image and choose where to save the PPT file. The tool will convert the image and notify you when the slide has been created.

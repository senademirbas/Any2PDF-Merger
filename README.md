# Any2PDF Merger

A Python-based desktop application designed to convert various file formats into a single, merged PDF document. Built with `tkinter` for the user interface.

## Supported Formats & Engine

- **Jupyter Notebook (`.ipynb`) & HTML**: Converted natively using `nbconvert` and headless Microsoft Edge PDF rendering, preserving cells, plots, and graphical layouts.
- **Images (PNG, JPG, BMP, etc.)**: Embedded directly into the PDF at their native resolution using `reportlab`, avoiding rasterization loss or compression artifacts.
- **Office Documents (DOCX, PPTX, XLSX)**: Converted silently via `comtypes` (requires Microsoft Office).
- **Text & Source Code**: Plain text and code files (.py, .cpp, .json) are rendered using a text-to-pdf engine parsing standard word wraps.

## Requirements

- Python 3.x
- **Python Packages**: `Pillow`, `pypdf`, `reportlab`, `nbformat`, `nbconvert`, `comtypes`
- **System**: Microsoft Edge (for `.ipynb`/HTML rendering) and Microsoft Office (for Office formats).

## Installation

1. Clone the repository:
   ```cmd
   git clone https://github.com/senademirbas/Any2PDF-Merger.git
   cd Any2PDF-Merger
   ```

2. Install the dependencies:
   ```cmd
   pip install Pillow pypdf reportlab nbformat nbconvert comtypes
   ```

## Usage

Simply run the application script in your terminal:

```cmd
python pdf_birlestirici.py
```

Use the interface to add files, adjust their merge order with the arrow buttons, and generate the final combined PDF document.

## Contributing

This project is completely open to suggestions, bug fixes, and improvements. Feel free to fork the repository, enhance the code, and submit pull requests. All contributions are welcome!

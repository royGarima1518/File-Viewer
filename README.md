# Smart File Reader & Editor Tool🔎

This is a Python-based GUI application built with Tkinter that allows users to open, edit, search, and replace content in a variety of file formats, including:

- Word Documents (`.docx`)
- Excel Sheets (`.xlsx`)
- PowerPoint Presentations (`.pptx`)
- Text files (`.txt`, `.py`, `.js`, `.html`, `.c`, `.h`)
- Email files (`.msg`, `.eml`)
- PDF files (`.pdf`)
- Image files (`.png`, `.jpg`, `.jpeg`, `.gif`)

## Features

- **Open Files**: Open multiple files at once.
- **View Files**: View the content of files in a tabbed interface.
- **Search and Replace**: Search for specific text and replace all occurrences.
- **Navigate Matches**: Navigate between the search matches.
- **Preview Images**: View images in their original size within the application.
- **Save Files**: Save the modified files with versioning support.
- **Excel Sheet Selection**: Choose specific sheets to view and edit for Excel files.

## Installation

To run this application, you need to have Python installed on your machine. You also need to install several Python libraries, which are listed below.

### Requirements

- Python 3.x
- Tkinter (usually comes pre-installed with Python)
- `python-docx` - For reading `.docx` files
- `openpyxl` - For reading and writing `.xlsx` files
- `python-pptx` - For reading `.pptx` files
- `extract-msg` - For reading email `.msg` files
- `PyMuPDF` (`fitz`) - For reading `.pdf` files
- `Pillow` - For handling image files

### Installing Required Libraries

To install the required libraries, run the following command:

```bash
pip install python-docx openpyxl python-pptx extract-msg PyMuPDF Pillow

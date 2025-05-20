# PDF to DOCX Converter

This project is a graphical interface for converting PDF files to DOCX format. The program supports text, table, and mathematical formula processing, and also uses OCR for extracting text from images in PDF files.

- **PyQt6**: for creating the graphical user interface (GUI)
- **pytesseract**: for OCR
- **pdfplumber**: for extracting text and tables from PDF files
- **PyMuPDF**: for working with PDF files
- **pikepdf**: for unlocking protected PDF files
- **pdf2docx**: for converting PDF to DOCX
- **python-docx**: for working with DOCX files
- **Pillow**: for image convertion

## How the Project Works

1. **Select a PDF File**: The user selects a PDF file for conversion
2. **Select an out folder**: The user specifies the folder where the result will be saved
3. **Convert**: The program performs the following steps:
   - If the PDF file is password-protected, the program unlocks it using the `pikepdf` library
   - If the PDF contains images, the program uses `pytesseract` for OCR.
   - Text, tables, and mathematical formulas are extracted using `pdfplumber` and `PyMuPDF`
   - The final result is saved in DOCX format using `python-docx`

## Installation and Launch

1. Ensure that Python version 3.10 or higher is installed on your system
 Install the dependencies from the requirements.txt file:
`pip install -r requirements.txt`
1. Or run `git clone https://github.com/15SAVAGE/PDF-to-DOCX.git`
then `pip install .`
3. Download and install Tesseract OCR from the official repository
https://github.com/tesseract-ocr/tesseract
After installation, copy the following components to your project folder (with this python script):
tesseract.exe
tessdata folder (you need to install it yourself)


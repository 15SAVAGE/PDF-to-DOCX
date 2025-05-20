from setuptools import setup, find_packages
setup(
    name="pdf_to_docx_converter",
    version="1.0.0",
    description="A graphical interface for converting PDF files to DOCX format with OCR support.",
    author="15SAVAGE",
    author_email="example@mail.com",
    url="https://github.com/15SAVAGE/PDF-to-DOCX",
    install_requires=[
        "pytesseract",
        "pdfplumber",
        "PyMuPDF",
        "pikepdf",
        "pdf2docx",
        "python-docx",
        "PyQt6",
        "Pillow"
    ],
    classifiers=[
        "Programming Language :: Python :: 3",
        "License :: OSI Approved :: MIT License",
        "Operating System :: OS Independent",
    ],
    python_requires=">=3.10",
)

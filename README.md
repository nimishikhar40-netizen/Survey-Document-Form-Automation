# Survey-Document-Form-Automation
A robust suite of Python automation tools designed to extract company contact details (Names, Phone Numbers, and Email Addresses) from scanned or digital PDF directories and export them directly into structured Excel (`.xlsx`) spreadsheets.
--

## 🚀 Key Features

*   **Multi-mode Extraction**:
    *   **Standard PDF Text Parsing**: Programmatic text extraction using `pdfplumber` for digital PDFs.
    *   **Scanned OCR Extraction**: Optical Character Recognition using `pytesseract` and `pdf2image` for scanned pages and images.
    *   **Pre-structured Markdown Exporter**: Direct serialization of markdown-formatted contact tables to Excel.
*   **Comprehensive Dependencies Installer**: Automated Poppler utility downloading for Windows.
*   **Clean Exporting**: Auto-generates clean headers (`Company Name`, `Contact Number`, `Email`) with built-in text trimming and sanitization.

---

## 📁 Project Architecture

| File | Description |
| :--- | :--- |
| [`extract_pdf_to_excel.py`](file:///c:/Users/user/Desktop/auto/extract_pdf_to_excel.py) | Parses inline markdown table entries and exports them to `company_contacts_from_markdown.xlsx`. |
| [`create_companies_excel.py`](file:///c:/Users/user/Desktop/auto/create_companies_excel.py) | Contains parsed data chunks (letters M-Z) and compiles them into `companies.xlsx`. |
| [`extract_all_companies.py`](file:///c:/Users/user/Desktop/auto/extract_all_companies.py) | Scans digital PDF files with `pdfplumber` to pattern-match company suffix structures. |
| [`extract_pdf_with_ocr.py`](file:///c:/Users/user/Desktop/auto/extract_pdf_with_ocr.py) | Runs OCR using Tesseract on image-based PDFs, page-by-page. |
| [`install_poppler.ps1`](file:///c:/Users/user/Desktop/auto/install_poppler.ps1) | PowerShell installer script that downloads, extracts, and places Poppler in the project folder. |
| [`run_ocr.bat`](file:///c:/Users/user/Desktop/auto/run_ocr.bat) | Batch command shortcut to start the OCR extraction script in one click. |

---

## 🛠️ Prerequisites & Setup

### 1. Python Environment
First, ensure you have Python 3.8+ installed. Set up your virtual environment and install the required libraries:

```bash
# Create a virtual environment
python -m venv .venv

# Activate the virtual environment
# On Windows:
.venv\Scripts\activate

# Install required packages
pip install openpyxl pandas pdfplumber pdf2image pytesseract
```

### 2. Install Poppler (Required for PDF-to-Image conversion)
For Windows users, Poppler is required to convert PDF pages to images. You can automate this installation using the included PowerShell script:

```powershell
powershell -ExecutionPolicy Bypass -File install_poppler.ps1
```
This will automatically download and extract Poppler into a local `poppler` directory within the project root.

### 3. Install Tesseract OCR (Required for Scanned PDFs)
To perform OCR on images or scanned documents:
1. Download the installer from the [Tesseract OCR for Windows](https://github.com/UB-Mannheim/tesseract/wiki) wiki.
2. Alternatively, install it using Windows Package Manager (`winget`):
   ```bash
   winget install UB-Mannheim.TesseractOCR
   ```
3. Make sure Tesseract is added to your system environment variables, or configured in `extract_pdf_with_ocr.py`.

---

## 💻 How to Use

### Run Markdown-to-Excel Compiler
To compile the pre-structured markdown table (containing Aaditya Stainless, Alfa Pumps, etc.) to Excel:
```bash
python extract_pdf_to_excel.py
```

### Run M-Z Company List Compiler
To compile the parsed dataset containing letters M-Z:
```bash
python create_companies_excel.py
```

### Run Programmatic PDF Extractor
If your PDF (`SKM_C284e25121822480_compressed.pdf`) has a readable text layer, parse it directly:
```bash
python extract_all_companies.py
```

### Run OCR Extractor (Scanned PDFs)
For scanned PDFs, run the OCR script:
```bash
python extract_pdf_with_ocr.py
```
Or use the automated launcher:
```cmd
run_ocr.bat
```

---

## 📊 Output Schema
The generated spreadsheets (`companies.xlsx` and `company_contacts_from_markdown.xlsx`) will follow this layout:

| Company Name | Contact Number | Email |
| :--- | :--- | :--- |
| AADITYA STAINLESS PVT. LTD. | 917021085650 | marketing@aadityastainless.com |
| AANURAJ FASTENERS PVT. LTD. | (Blank/Empty) | sales@aanuraj.com |
| NETZSCH TECHNOLOGIES INDIA PVT. LTD. | 4442965100 | prasanna.samanthula@netzsch.com |


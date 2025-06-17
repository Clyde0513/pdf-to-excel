# PDF to Excel Converter with Registry Staff Parser

This script converts scanned PDF files to Excel format using OCR (Optical Character Recognition) with specialized parsing for Registry Staff documents. It automatically extracts staff lists organized by categories and creates a structured Excel output.

## Features

- **Smart OCR Processing**: Automatically detects text-based vs scanned PDFs using Tesseract
- **Registry Staff Parser**: Specialized parsing for documents with staff categories (TheraEX, Intuitive, Vitawerks, Vitawerks Cont)
- **Structured Excel Output**: Creates organized columns for each staff category with proper alignment
- **Auto-Detection**: Automatically finds and configures Tesseract OCR installation
- **Summary Statistics**: Provides staff counts and conversion metrics
- **Multiple Sheets**: Raw text, parsed data, and summary information
- **Complete Data Capture**: Ensures all staff names and credentials are extracted and included

## Prerequisites

### Install Tesseract OCR

**Option 1: Automated Installation (Recommended)**
```powershell
# Run the automated installer (as Administrator for best results)
powershell -ExecutionPolicy Bypass -File install_tesseract_fixed.ps1
```

**Option 2: Manual Installation**
1. Download: [Tesseract 5.5.0 Direct Link](https://github.com/tesseract-ocr/tesseract/releases/download/5.5.0/tesseract-ocr-w64-setup-5.5.0.20241111.exe)
2. Run installer as Administrator
3. Install to default directory: `C:\Program Files\Tesseract-OCR\`

**Option 3: Quick Download**
```powershell
# Download installer
Invoke-WebRequest -Uri "https://github.com/tesseract-ocr/tesseract/releases/download/5.5.0/tesseract-ocr-w64-setup-5.5.0.20241111.exe" -OutFile "tesseract-installer.exe"

# Run installer
Start-Process -FilePath "tesseract-installer.exe" -Wait
```

## Installation

1. **Install Python packages:**
   ```powershell
   pip install -r requirements.txt
   ```

2. **Verify Tesseract installation:**
   ```powershell
   tesseract --version
   ```

## Usage

### Basic Conversion
```powershell
# Convert default PDF file
python pdf_to_excel_pymupdf.py

# Convert specific file
python pdf_to_excel_pymupdf.py "your-document.pdf"

# Specify output filename
python pdf_to_excel_pymupdf.py "input.pdf" "output.xlsx"
```

### Registry Staff Documents
For documents containing registry staff lists with categories like:
- **TheraEX**
- **Intuitive** 
- **Vitawerks**
- **Vitawerks Cont**

The script automatically creates separate columns for each category with staff names and credentials.

### Debug and Verification
```powershell
# See extracted text and parsing analysis
python debug_pdf.py

# View complete Excel output
python show_all_rows.py

# Check file structure
python show_excel.py "output.xlsx"
```

## Output Format

The Excel file contains multiple sheets with comprehensive data:

### Main Data Sheet
- **Category Columns**: Each staff category (TheraEX, Intuitive, Vitawerks, Vitawerks Cont) gets its own column
- **Staff Names**: Names with credentials listed under the appropriate category
- **Variable Length Lists**: Handles different numbers of staff per category with proper alignment
- **Complete Coverage**: All names are captured and included in the output

### Additional Sheets
- **Raw Text**: Complete OCR extracted text for reference
- **Summary**: Statistics including staff counts per category and conversion metrics

### Data Structure
For registry staff documents, the output automatically organizes data into:
- Column headers matching staff categories found in the PDF
- Staff names and credentials aligned under correct categories  
- Empty cells where categories have fewer staff members
- Page numbers and source line references for verification

## Verification

The tool includes several verification scripts to ensure data completeness:

```powershell
# View all extracted data (not just first 10 rows)
python show_all_rows.py

# See detailed parsing analysis
python detailed_debug.py

# Check Excel file structure
python show_excel.py "Document250616132824_v3.xlsx"
```

**Note**: The Excel output contains ALL extracted names, not just a preview. Use the verification scripts above to confirm all data is present.

## Project Status

This tool has been **successfully tested and verified** with registry staff documents. Key accomplishments:

- **Complete Data Extraction**: All staff names and credentials are captured
- **Proper Categorization**: Staff correctly organized under category headers
- **Variable List Handling**: Different numbers of staff per category handled correctly
- **Quality Verification**: Multiple verification scripts confirm data completeness
- **User-Friendly Setup**: Automated installation and clear instructions provided

### Sample Results
The tool successfully processes documents like `Document250616132824.pdf` and generates structured Excel files (`Document250616132824_v3.xlsx`) with:
- TheraEX staff in column 1
- Intuitive staff in column 2  
- Vitawerks staff in column 3
- Vitawerks Cont staff in column 4
- All names with credentials preserved
- Summary statistics and raw text for reference

## Customization

You may need to modify the `parse_text_to_structured_data` method based on your specific PDF format. The current implementation:
- Detects tabular data by looking for multiple spaces or tabs
- Splits data into columns accordingly
- Handles single-column data as well

## Troubleshooting

### Common Issues

1. **Tesseract not found**: 
   - Use the automated installer: `powershell -ExecutionPolicy Bypass -File install_tesseract_fixed.ps1`
   - Or manually install to `C:\Program Files\Tesseract-OCR\`

2. **Poor OCR results**: 
   - Try increasing the DPI in the script settings
   - Ensure the PDF image quality is good

3. **Memory issues**: 
   - For large PDFs, the script processes pages individually
   - Close other applications if memory is limited

4. **Missing dependencies**: 
   - Run `pip install -r requirements.txt` to install all required packages

5. **Excel file not opening**: 
   - Check that the output file isn't already open in Excel
   - Verify file permissions in the output directory

### Additional Help Files

The repository includes several helper documents:
- `TESSERACT_INSTALL.md` - Detailed Tesseract installation guide
- `debug_output.txt` - Sample debugging output for reference
- Multiple verification scripts for data checking

## Files Overview

### Main Scripts
- `pdf_to_excel_pymupdf.py` - Main conversion script with advanced registry staff parsing
- `requirements.txt` - Python package dependencies

### Installation Helpers  
- `install_tesseract_fixed.ps1` - Automated Tesseract installer
- `setup.ps1` - Complete environment setup script
- `TESSERACT_INSTALL.md` - Manual installation guide

### Verification Tools
- `show_all_rows.py` - Display complete Excel output data
- `detailed_debug.py` - Comprehensive parsing analysis  
- `debug_pdf.py` - OCR extraction debugging
- `show_excel.py` - Excel file structure viewer
- `show_complete_data.py` - Complete data verification

## Dependencies

### Python Packages (installed via requirements.txt)
- `pytesseract` - OCR engine interface for text extraction
- `pdf2image` - PDF to image conversion for OCR processing  
- `pandas` - Data manipulation and Excel file creation
- `openpyxl` - Excel file writing and formatting
- `Pillow (PIL)` - Image processing and enhancement
- `numpy` - Numerical operations and data handling
- `PyMuPDF (fitz)` - PDF processing and text extraction

### External Dependencies
- **Tesseract OCR** - Text recognition engine (automatically detected and configured)
- **Poppler** - PDF rendering (included with pdf2image package)

### Installation
All Python dependencies are automatically installed with:
```powershell
pip install -r requirements.txt
```

The script automatically detects and configures Tesseract OCR installation.

## Quick Start Guide

1. **Install Tesseract**: Run `powershell -ExecutionPolicy Bypass -File install_tesseract_fixed.ps1`
2. **Install Python packages**: `pip install -r requirements.txt`  
3. **Convert your PDF**: `python pdf_to_excel_pymupdf.py "your-document.pdf"`
4. **Verify results**: `python show_all_rows.py` to see all extracted data
---
import pytesseract
import pandas as pd
import fitz  # PyMuPDF
from PIL import Image
import re
import os
import sys
import io

# Debug version to see extracted text
class DebugPDFToExcelConverter:
    def __init__(self, pdf_path):
        self.pdf_path = pdf_path
        # Auto-detect Tesseract path
        self._setup_tesseract_path()
    
    def _setup_tesseract_path(self):
        import shutil
        import subprocess
        
        tesseract_cmd = shutil.which('tesseract')
        
        if tesseract_cmd:
            print(f"Found Tesseract in PATH: {tesseract_cmd}")
            return
        
        possible_paths = [
            r'C:\Program Files\Tesseract-OCR\tesseract.exe',
            r'C:\Program Files (x86)\Tesseract-OCR\tesseract.exe',
            r'C:\Tesseract-OCR\tesseract.exe',
        ]
        
        for path in possible_paths:
            if os.path.exists(path):
                print(f"Found Tesseract at: {path}")
                pytesseract.pytesseract.tesseract_cmd = path
                
                try:
                    version = subprocess.check_output([path, '--version'], 
                                                    stderr=subprocess.STDOUT, 
                                                    universal_newlines=True)
                    print(f"Tesseract version: {version.splitlines()[0]}")
                    return
                except Exception as e:
                    print(f"Tesseract found but not working: {e}")
                    continue
        
        raise FileNotFoundError("Tesseract OCR not found. Please install it first.")
    
    def debug_extract_text(self):
        try:
            print("Opening PDF...")
            doc = fitz.open(self.pdf_path)
            
            for page_num in range(len(doc)):
                print(f"\n--- PAGE {page_num + 1} ---")
                page = doc.load_page(page_num)
                
                # Try direct text extraction first
                text = page.get_text()
                
                if text.strip():
                    print("DIRECT TEXT FOUND:")
                    print(repr(text))
                else:
                    print("No direct text, using OCR...")
                    
                    # Convert page to image and OCR
                    mat = fitz.Matrix(2, 2)
                    pix = page.get_pixmap(matrix=mat)
                    img_data = pix.tobytes("png")
                    
                    img = Image.open(io.BytesIO(img_data))
                    ocr_text = pytesseract.image_to_string(img, lang='eng')
                    
                    print("OCR TEXT:")
                    print(repr(ocr_text))
                    
                    print("\nOCR TEXT (formatted):")
                    print(ocr_text)
                    
                    # Try to parse the registry staff structure
                    print("\n--- PARSING ANALYSIS ---")
                    self._analyze_registry_structure(ocr_text)
            
            doc.close()
            
        except Exception as e:
            print(f"Error: {str(e)}")
    
    def _analyze_registry_structure(self, text):
        lines = [line.strip() for line in text.split('\n') if line.strip()]
        
        print(f"Total lines: {len(lines)}")
        
        # Look for registry staff header
        registry_start = -1
        for i, line in enumerate(lines):
            if 'Registry Staff' in line and 'not in Qgenda' in line:
                registry_start = i
                print(f"Found 'Registry Staff' header at line {i}: {repr(line)}")
                break
        
        if registry_start == -1:
            print("No 'Registry Staff' header found!")
            print("All lines:")
            for i, line in enumerate(lines):
                print(f"  {i}: {repr(line)}")
            return
        
        # Look for headers
        headers = ['TheraEX', 'Intuitive', 'Vitawerks', 'Vitawerks Cont']
        
        print(f"\nAnalyzing lines after registry header (starting from line {registry_start + 1}):")
        current_header = None
        
        for i in range(registry_start + 1, len(lines)):
            line = lines[i]
            print(f"  {i}: {repr(line)}")
            
            # Check if this is a header
            if line.endswith(':'):
                header_name = line.rstrip(':')
                if header_name in headers:
                    current_header = header_name
                    print(f"    -> HEADER: {current_header}")
                    continue
            
            # Check for separator
            if re.match(r'^_+$', line):
                print(f"    -> SEPARATOR")
                continue
            
            # Check if looks like a name
            if self._looks_like_staff_name(line):
                print(f"    -> STAFF NAME under {current_header}: {line}")
            else:
                print(f"    -> OTHER: {line}")
    
    def _looks_like_staff_name(self, line):
        if re.search(r'[A-Za-z]+\s+[A-Za-z]+,?\s*[A-Z]{2,4}$', line):
            return True
        
        credentials = ['RN', 'MD', 'NP', 'PA', 'LPN', 'CNA', 'RT', 'RPh', 'PharmD', 'DPT', 'OT', 'PT', 'SLP']
        for cred in credentials:
            if line.endswith(f', {cred}') or line.endswith(f' {cred}'):
                return True
        
        words = line.split()
        if len(words) >= 2 and all(word.replace(',', '').isalpha() for word in words[:2]):
            return True
        
        return False

# Run debug
if __name__ == "__main__":
    pdf_path = "Document250616132824.pdf"
    if len(sys.argv) > 1:
        pdf_path = sys.argv[1]
    
    converter = DebugPDFToExcelConverter(pdf_path)
    converter.debug_extract_text()

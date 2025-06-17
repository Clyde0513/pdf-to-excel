import pytesseract
import pandas as pd
import fitz
from PIL import Image
import re
import os
import io

# Check what data structure is being created
class DetailedDebugConverter:
    def __init__(self, pdf_path):
        self.pdf_path = pdf_path
        self._setup_tesseract_path()
    
    def _setup_tesseract_path(self):
        import shutil
        import subprocess
        
        tesseract_cmd = shutil.which('tesseract')
        if tesseract_cmd:
            return
        
        possible_paths = [
            r'C:\Program Files\Tesseract-OCR\tesseract.exe',
            r'C:\Program Files (x86)\Tesseract-OCR\tesseract.exe',
        ]
        
        for path in possible_paths:
            if os.path.exists(path):
                pytesseract.pytesseract.tesseract_cmd = path
                return
    
    def debug_data_structure(self):
        # Extract text
        doc = fitz.open(self.pdf_path)
        page = doc.load_page(0)
        
        text = page.get_text()
        if not text.strip():
            mat = fitz.Matrix(2, 2)
            pix = page.get_pixmap(matrix=mat)
            img_data = pix.tobytes("png")
            img = Image.open(io.BytesIO(img_data))
            text = pytesseract.image_to_string(img, lang='eng')
        
        doc.close()
        
        # Parse exactly like the main script
        headers = ['TheraEX', 'Intuitive', 'Vitawerks', 'Vitawerks Cont']
        lines = [line.strip() for line in text.split('\n') if line.strip()]
        
        # Find registry staff section
        registry_start = -1
        for i, line in enumerate(lines):
            if 'Registry Staff' in line and 'not in Qgenda' in line:
                registry_start = i
                break
        
        # Initialize data structure
        staff_data = {header: [] for header in headers}
        current_header = None
        
        print("=== DETAILED PARSING ===")
        
        # Process lines after the registry staff header
        for i in range(registry_start + 1, len(lines)):
            line = lines[i]
            
            # Check if this line is a header (ends with colon)
            if line.endswith(':'):
                header_name = line.rstrip(':')
                if header_name in headers:
                    current_header = header_name
                    print(f"\n🏷️  HEADER: {current_header}")
                    continue
            
            # Skip separators and short lines
            if re.match(r'^_+$', line) or len(line) < 3:
                continue
            
            # If we have a current header and this looks like a name
            if current_header and self._looks_like_staff_name(line):
                staff_data[current_header].append(line)
                print(f"   {len(staff_data[current_header])}: {line}")
        
        print("\n=== FINAL COUNTS ===")
        for header in headers:
            count = len(staff_data[header])
            print(f"{header}: {count} people")
            if count > 0:
                print(f"  First: {staff_data[header][0]}")
                print(f"  Last:  {staff_data[header][-1]}")
        
        # Show how the Excel structure will be created
        max_length = max(len(names) for names in staff_data.values()) if any(staff_data.values()) else 0
        print(f"\n=== EXCEL STRUCTURE ===")
        print(f"Max length (rows): {max_length}")
        
        print(f"\nFirst 5 rows will be:")
        for row_idx in range(min(5, max_length)):
            row_data = {'Row': row_idx + 1}
            for header in headers:
                if row_idx < len(staff_data[header]):
                    row_data[header] = staff_data[header][row_idx]
                else:
                    row_data[header] = ''
            print(f"  Row {row_idx + 1}: {row_data}")
        
        print(f"\nLast 5 rows will be:")
        start_idx = max(0, max_length - 5)
        for row_idx in range(start_idx, max_length):
            row_data = {'Row': row_idx + 1}
            for header in headers:
                if row_idx < len(staff_data[header]):
                    row_data[header] = staff_data[header][row_idx]
                else:
                    row_data[header] = ''
            print(f"  Row {row_idx + 1}: {row_data}")
        
        return staff_data
    
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

if __name__ == "__main__":
    converter = DetailedDebugConverter("Document250616132824.pdf")
    staff_data = converter.debug_data_structure()

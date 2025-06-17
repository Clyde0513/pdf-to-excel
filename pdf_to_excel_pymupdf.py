import pytesseract
import pandas as pd
import fitz  # PyMuPDF
from PIL import Image
import re
import os
import sys
import io

class PDFToExcelConverter:
    def __init__(self, pdf_path, output_path=None):
        """
        Initialize the PDF to Excel converter.
        
        Args:
            pdf_path (str): Path to the PDF file
            output_path (str): Path for the output Excel file (optional)
        """
        self.pdf_path = pdf_path
        self.output_path = output_path or pdf_path.replace('.pdf', '.xlsx')
          # Auto-detect Tesseract path on Windows
        self._setup_tesseract_path()
    
    def _setup_tesseract_path(self):
        """
        Auto-detect and set up Tesseract path on Windows.
        """
        import shutil
        import subprocess
        
        # Try to find tesseract in PATH first
        tesseract_cmd = shutil.which('tesseract')
        
        if tesseract_cmd:
            print(f"Found Tesseract in PATH: {tesseract_cmd}")
            return
        
        # Common Windows installation paths
        possible_paths = [
            r'C:\Program Files\Tesseract-OCR\tesseract.exe',
            r'C:\Program Files (x86)\Tesseract-OCR\tesseract.exe',
            r'C:\Tesseract-OCR\tesseract.exe',
        ]
        
        for path in possible_paths:
            if os.path.exists(path):
                print(f"Found Tesseract at: {path}")
                pytesseract.pytesseract.tesseract_cmd = path
                
                # Test if it works
                try:
                    version = subprocess.check_output([path, '--version'], 
                                                    stderr=subprocess.STDOUT, 
                                                    universal_newlines=True)
                    print(f"Tesseract version: {version.splitlines()[0]}")
                    return
                except Exception as e:
                    print(f"Tesseract found but not working: {e}")
                    continue
        
        # If we get here, Tesseract wasn't found
        print("Tesseract OCR not found!")
        print("Please install Tesseract OCR:")
        print("1. Run: .\\install_tesseract.ps1")
        print("2. Or download from: https://github.com/UB-Mannheim/tesseract/wiki")
        print("3. Or install using: choco install tesseract")
        raise FileNotFoundError("Tesseract OCR not found. Please install it first.")
    
    def extract_text_from_pdf(self):
        """
        Extract text from PDF using PyMuPDF and OCR.
        
        Returns:
            list: List of text content from each page
        """
        try:
            # Open PDF with PyMuPDF
            print("Opening PDF...")
            doc = fitz.open(self.pdf_path)
            
            extracted_text = []
            
            for page_num in range(len(doc)):
                print(f"Processing page {page_num + 1}/{len(doc)}...")
                page = doc.load_page(page_num)
                
                # First try to extract text directly (for text-based PDFs)
                text = page.get_text()
                
                if text.strip():
                    print(f"Found text directly on page {page_num + 1}")
                    extracted_text.append(text)
                else:
                    # If no text found, use OCR on the page image
                    print(f"No direct text found, using OCR on page {page_num + 1}")
                    
                    # Convert page to image
                    mat = fitz.Matrix(2, 2)  # 2x zoom for better OCR
                    pix = page.get_pixmap(matrix=mat)
                    img_data = pix.tobytes("png")
                    
                    # Convert to PIL Image
                    img = Image.open(io.BytesIO(img_data))
                    
                    # Perform OCR on the image
                    ocr_text = pytesseract.image_to_string(img, lang='eng')
                    extracted_text.append(ocr_text)
            
            doc.close()
            return extracted_text
            
        except Exception as e:
            print(f"Error extracting text from PDF: {str(e)}")
            return []
    
    def parse_text_to_structured_data(self, text_pages):
        """
        Parse extracted text into structured data for Registry Staff format.
        
        Args:
            text_pages (list): List of text content from each page
        
        Returns:
            list: List of dictionaries representing rows of data
        """
        all_data = []
        
        for page_num, text in enumerate(text_pages):
            if not text.strip():
                continue
            
            # Parse registry staff data
            registry_data = self._parse_registry_staff(text, page_num)
            if registry_data:
                all_data.extend(registry_data)
            else:
                # Fallback to generic parsing if registry parsing doesn't work
                generic_data = self._generic_parse(text, page_num)
                all_data.extend(generic_data)
        
        return all_data
    
    def _parse_registry_staff(self, text, page_num):
        """
        Parse registry staff data with specific headers like TheraEX, Intuitive, etc.
        
        Args:
            text (str): Text content from the page
            page_num (int): Page number
        
        Returns:
            list: List of dictionaries with registry staff data
        """
        # Define expected headers
        headers = ['TheraEX', 'Intuitive', 'Vitawerks', 'Vitawerks Cont']
        
        lines = [line.strip() for line in text.split('\n') if line.strip()]
        
        # Find registry staff section
        registry_start = -1
        for i, line in enumerate(lines):
            if 'Registry Staff' in line and 'not in Qgenda' in line:
                registry_start = i
                break
        
        if registry_start == -1:
            return []
        
        # Initialize data structure
        staff_data = {header: [] for header in headers}
        current_header = None
        
        # Process lines after the registry staff header
        for i in range(registry_start + 1, len(lines)):
            line = lines[i]
            
            # Check if this line is a header (ends with colon)
            if line.endswith(':'):
                header_name = line.rstrip(':')
                if header_name in headers:
                    current_header = header_name
                    continue
            
            # Check if this is a separator line (underscores)
            if re.match(r'^_+$', line):
                continue
            
            # Skip empty or very short lines
            if len(line) < 3:
                continue
            
            # If we have a current header and this looks like a name
            if current_header and self._looks_like_staff_name(line):
                staff_data[current_header].append(line)
        
        # Convert to list of dictionaries for Excel
        result = []
        max_length = max(len(names) for names in staff_data.values()) if any(staff_data.values()) else 0
        
        for row_idx in range(max_length):
            row_data = {'Row_Number': row_idx + 1, 'Page': page_num + 1}
            for header in headers:
                if row_idx < len(staff_data[header]):
                    row_data[header] = staff_data[header][row_idx]
                else:
                    row_data[header] = ''
            result.append(row_data)
        
        return result
    
    def _looks_like_staff_name(self, line):
        """
        Check if a line looks like a staff name with credentials.
        
        Args:
            line (str): Line to check
        
        Returns:
            bool: True if it looks like a staff name
        """
        # Look for patterns like "Name, RN" or "Name Name, MD" etc.
        if re.search(r'[A-Za-z]+\s+[A-Za-z]+,?\s*[A-Z]{2,4}$', line):
            return True
        
        # Look for names with common credentials
        credentials = ['RN', 'MD', 'NP', 'PA', 'LPN', 'CNA', 'RT', 'RPh', 'PharmD', 'DPT', 'OT', 'PT', 'SLP']
        for cred in credentials:
            if line.endswith(f', {cred}') or line.endswith(f' {cred}'):
                return True
        
        # Look for two or more words (likely a name)
        words = line.split()
        if len(words) >= 2 and all(word.replace(',', '').isalpha() for word in words[:2]):
            return True
        
        return False
    
    def _generic_parse(self, text, page_num):
        """
        Generic fallback parsing method.
        
        Args:
            text (str): Text content
            page_num (int): Page number
        
        Returns:
            list: List of dictionaries
        """
        lines = [line.strip() for line in text.split('\n') if line.strip()]
        result = []
        
        for line in lines:
            if len(line) < 3:
                continue
            
            row_data = {
                'Content': line,
                'Page': page_num + 1,
                'Source_Line': line
            }
            result.append(row_data)
        
        return result
    
    def detect_table_structure(self, text_pages):
        """
        Advanced table detection and parsing.
        
        Args:
            text_pages (list): List of text content from each page
        
        Returns:
            list: List of dictionaries representing table rows
        """
        all_tables = []
        
        for page_num, text in enumerate(text_pages):
            if not text.strip():
                continue
            
            lines = [line.strip() for line in text.split('\n') if line.strip()]
            
            # Try to detect header row (usually contains multiple words separated by spaces)
            header_candidates = []
            for i, line in enumerate(lines[:10]):  # Check first 10 lines for headers
                if re.search(r'[A-Za-z].*\s+[A-Za-z]', line) and not re.search(r'^\d', line):
                    header_candidates.append((i, line))
            
            # Process potential table data
            for line in lines:
                # Look for lines with numbers and text that might be table rows
                if re.search(r'\d', line) and len(line.split()) > 1:
                    # Split by multiple spaces or common delimiters
                    parts = re.split(r'\s{2,}|\t|[|;,]', line)
                    parts = [part.strip() for part in parts if part.strip()]
                    
                    if len(parts) >= 2:
                        row_data = {
                            'Page': page_num + 1,
                            'Row_Type': 'Data'
                        }
                        
                        for i, part in enumerate(parts):
                            row_data[f'Column_{i+1}'] = part
                        
                        all_tables.append(row_data)
        
        return all_tables
    
    def create_excel_file(self, data):
        """
        Create Excel file from structured data.
        
        Args:
            data (list): List of dictionaries representing rows of data
        """
        if not data:
            print("No data to write to Excel file.")
            return
        
        try:
            # Create DataFrame
            df = pd.DataFrame(data)
              # Write to Excel with multiple sheets
            with pd.ExcelWriter(self.output_path, engine='openpyxl') as writer:
                # Write main data
                df.to_excel(writer, sheet_name='Registry_Staff', index=False)
                
                # Create a summary sheet
                registry_headers = ['TheraEX', 'Intuitive', 'Vitawerks', 'Vitawerks Cont']
                has_registry_data = any(col in df.columns for col in registry_headers)
                
                if has_registry_data:
                    # Registry-specific summary
                    summary_data = {
                        'Category': ['Total_Rows', 'Total_Pages', 'Source_File'] + registry_headers,
                        'Value': [
                            len(df),
                            df['Page'].nunique() if 'Page' in df.columns else 1,
                            os.path.basename(self.pdf_path)
                        ] + [
                            len(df[df[header].notna() & (df[header] != '')]) if header in df.columns else 0
                            for header in registry_headers
                        ]
                    }
                else:
                    # Generic summary
                    summary_data = {
                        'Metric': ['Total_Rows', 'Total_Pages', 'Total_Columns', 'Source_File'],
                        'Value': [
                            len(df),
                            df['Page'].nunique() if 'Page' in df.columns else 1,
                            len([col for col in df.columns if col.startswith('Column_')]),
                            os.path.basename(self.pdf_path)
                        ]
                    }
                summary_df = pd.DataFrame(summary_data)
                summary_df.to_excel(writer, sheet_name='Summary', index=False)
                
                # Create a raw text sheet for reference
                if hasattr(self, '_raw_text'):
                    raw_text_data = []
                    for i, text in enumerate(self._raw_text):
                        raw_text_data.append({
                            'Page': i + 1,
                            'Raw_Text': text[:1000] + '...' if len(text) > 1000 else text
                        })
                    raw_df = pd.DataFrame(raw_text_data)
                    raw_df.to_excel(writer, sheet_name='Raw_Text', index=False)
            
            print(f"Excel file created successfully: {self.output_path}")
            print(f"Total rows extracted: {len(df)}")
            
            # Print column summary
            columns = [col for col in df.columns if col.startswith('Column_')]
            if columns:
                print(f"Data organized into {len(columns)} columns")
            
        except Exception as e:
            print(f"Error creating Excel file: {str(e)}")
    
    def convert(self):
        """
        Main conversion method.
        """
        print(f"Starting conversion of {self.pdf_path}...")
        
        # Check if PDF file exists
        if not os.path.exists(self.pdf_path):
            print(f"Error: PDF file not found at {self.pdf_path}")
            return
        
        # Extract text from PDF
        text_pages = self.extract_text_from_pdf()
        
        if not text_pages:
            print("No text could be extracted from the PDF.")
            return
        
        # Store raw text for reference
        self._raw_text = text_pages
        
        # Parse text to structured data
        print("Parsing extracted text...")
        structured_data = self.parse_text_to_structured_data(text_pages)
        
        # Also try advanced table detection
        table_data = self.detect_table_structure(text_pages)
        
        # Combine data (prefer table data if found)
        final_data = table_data if table_data else structured_data
        
        if not final_data:
            print("No structured data could be extracted. Creating a simple text dump...")
            final_data = []
            for i, text in enumerate(text_pages):
                final_data.append({
                    'Page': i + 1,
                    'Content': text
                })
        
        # Create Excel file
        print("Creating Excel file...")
        self.create_excel_file(final_data)
        
        print("Conversion completed!")

def main():
    """
    Main function to run the PDF to Excel converter.
    """
    # Get PDF file path from command line argument or use default
    if len(sys.argv) > 1:
        pdf_path = sys.argv[1]
    else:
        # Use the PDF file in the current directory
        pdf_path = "Document250616132824.pdf"
    
    # Get output path from command line argument (optional)
    output_path = sys.argv[2] if len(sys.argv) > 2 else None
    
    # Create converter instance
    converter = PDFToExcelConverter(pdf_path, output_path)
    
    # Run conversion
    converter.convert()

if __name__ == "__main__":
    main()

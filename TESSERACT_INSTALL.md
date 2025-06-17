# Tesseract OCR Installation Guide for Windows

## Manual Installation

### Step 1: Download Tesseract
1. Go to: https://github.com/tesseract-ocr/tesseract/releases/download/5.5.0/tesseract-ocr-w64-setup-5.5.0.20241111.exe
2. Or visit: https://github.com/UB-Mannheim/tesseract/wiki
3. Download the latest 64-bit installer: `tesseract-ocr-w64-setup-5.5.0.20241111.exe`

### Step 2: Run the Installer
1. Double-click the downloaded `.exe` file
2. **IMPORTANT**: Install to the default directory `C:\Program Files\Tesseract-OCR\` 
   - Do NOT install to an existing directory as the uninstaller removes the entire folder
3. During installation, make sure to:
   - Select "Additional language data" if you need languages other than English
   - Check the option to add Tesseract to PATH (if available)

### Step 3: Verify Installation
Open Command Prompt or PowerShell and run:
```powershell
tesseract --version
```

If this doesn't work, Tesseract is not in your PATH.

### Step 4: Add to PATH (if needed)
If `tesseract --version` doesn't work:

1. Open System Properties:
   - Press `Win + R`, type `sysdm.cpl`, press Enter
   - OR right-click "This PC" → Properties → Advanced system settings

2. Click "Environment Variables"

3. Under "System Variables", find and select "Path", then click "Edit"

4. Click "New" and add: `C:\Program Files\Tesseract-OCR`

5. Click "OK" on all dialogs

6. **Restart your terminal/PowerShell** and test again:
   ```powershell
   tesseract --version
   ```

## Automated Installation Script

Alternatively, you can use this PowerShell script to download and install automatically:

```powershell
# Download and install Tesseract automatically
$url = "https://github.com/tesseract-ocr/tesseract/releases/download/5.5.0/tesseract-ocr-w64-setup-5.5.0.20241111.exe"
$output = "$env:TEMP\tesseract-installer.exe"

Write-Host "Downloading Tesseract installer..." -ForegroundColor Yellow
Invoke-WebRequest -Uri $url -OutFile $output

Write-Host "Starting installation..." -ForegroundColor Yellow
Write-Host "Please follow the installation wizard." -ForegroundColor Green
Start-Process -FilePath $output -Wait

Write-Host "Installation completed!" -ForegroundColor Green
Write-Host "Please restart your terminal and run 'tesseract --version' to verify." -ForegroundColor Cyan
```

## After Installation

1. **Restart your terminal/PowerShell**
2. Test the installation:
   ```powershell
   tesseract --version
   ```
3. If successful, you can now run the PDF to Excel script:
   ```powershell
   python pdf_to_excel_pymupdf.py
   ```

## Troubleshooting

### If tesseract command is not found:
- Make sure you installed to the default directory
- Add `C:\Program Files\Tesseract-OCR` to your PATH environment variable
- Restart your terminal after adding to PATH

### If you get permission errors:
- Run the installer as Administrator
- Make sure you have write permissions to `C:\Program Files`

### Alternative: Manual PATH in Python script
If you can't add Tesseract to PATH, you can specify the path directly in the Python script by uncommenting and modifying this line:

```python
pytesseract.pytesseract.tesseract_cmd = r'C:\Program Files\Tesseract-OCR\tesseract.exe'
```

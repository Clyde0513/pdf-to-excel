# Setup script for PDF to Excel Converter

Write-Host "Setting up PDF to Excel Converter..." -ForegroundColor Green

# Check if Python is installed
try {
    $pythonVersion = python --version 2>&1
    Write-Host "Python found: $pythonVersion" -ForegroundColor Green
} catch {
    Write-Host "Python not found. Please install Python first." -ForegroundColor Red
    exit 1
}

# Install Python packages
Write-Host "Installing Python packages..." -ForegroundColor Yellow
pip install -r requirements.txt

if ($LASTEXITCODE -eq 0) {
    Write-Host "Python packages installed successfully!" -ForegroundColor Green
} else {
    Write-Host "Failed to install Python packages." -ForegroundColor Red
    exit 1
}

# Check if Tesseract is available
Write-Host "Checking for Tesseract OCR..." -ForegroundColor Yellow
try {
    $tesseractVersion = tesseract --version 2>&1
    Write-Host "Tesseract found: $($tesseractVersion[0])" -ForegroundColor Green
} catch {
    Write-Host "Tesseract OCR not found!" -ForegroundColor Red
    Write-Host "Please install Tesseract OCR from: https://github.com/UB-Mannheim/tesseract/wiki" -ForegroundColor Yellow
    Write-Host "After installation, you may need to add it to your PATH or update the script." -ForegroundColor Yellow
}

Write-Host "`nSetup completed!" -ForegroundColor Green
Write-Host "You can now run: python pdf_to_excel.py" -ForegroundColor Cyan

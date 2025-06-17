# Automated Tesseract Installation Script
# Run this script as Administrator for best results

Write-Host "=== Tesseract OCR Installation Script ===" -ForegroundColor Cyan
Write-Host ""

# Check if running as Administrator
$currentUser = [Security.Principal.WindowsIdentity]::GetCurrent()
$principal = New-Object Security.Principal.WindowsPrincipal($currentUser)
$isAdmin = $principal.IsInRole([Security.Principal.WindowsBuiltInRole]::Administrator)

if (-not $isAdmin) {
    Write-Host "WARNING: Not running as Administrator. Some steps may fail." -ForegroundColor Yellow
    Write-Host "For best results, right-click PowerShell and 'Run as Administrator'" -ForegroundColor Yellow
    Write-Host ""
}

# Download URL and local path
$tesseractUrl = "https://github.com/tesseract-ocr/tesseract/releases/download/5.5.0/tesseract-ocr-w64-setup-5.5.0.20241111.exe"
$installerPath = "$env:TEMP\tesseract-installer.exe"

try {
    # Download the installer
    Write-Host "Downloading Tesseract 5.5.0 installer..." -ForegroundColor Yellow
    Write-Host "From: $tesseractUrl" -ForegroundColor Gray
    Write-Host "To: $installerPath" -ForegroundColor Gray
    
    Invoke-WebRequest -Uri $tesseractUrl -OutFile $installerPath -UseBasicParsing
    
    if (Test-Path $installerPath) {
        $fileSize = (Get-Item $installerPath).Length / 1MB
        Write-Host "Download completed! File size: $([math]::Round($fileSize, 2)) MB" -ForegroundColor Green
    } else {
        throw "Download failed - installer file not found"
    }
    
    Write-Host ""
    Write-Host "Starting Tesseract installation..." -ForegroundColor Yellow
    Write-Host "IMPORTANT NOTES:" -ForegroundColor Red
    Write-Host "- Install to the default directory: C:\Program Files\Tesseract-OCR\" -ForegroundColor Red
    Write-Host "- Select additional language data if needed" -ForegroundColor Yellow
    Write-Host "- The installer window will open now..." -ForegroundColor Yellow
    Write-Host ""
    
    # Start the installer and wait for completion
    Start-Process -FilePath $installerPath -Wait
    
    Write-Host "Installation completed!" -ForegroundColor Green
    Write-Host ""
    
    # Check if Tesseract is installed
    Write-Host "Verifying installation..." -ForegroundColor Yellow
    
    $tesseractPath = "C:\Program Files\Tesseract-OCR\tesseract.exe"
    if (Test-Path $tesseractPath) {
        Write-Host "✓ Tesseract executable found at: $tesseractPath" -ForegroundColor Green
        
        # Check if it's in PATH
        try {
            $version = & tesseract --version 2>&1
            Write-Host "✓ Tesseract is in PATH and working!" -ForegroundColor Green
            Write-Host "Version info:" -ForegroundColor Gray
            Write-Host $version[0] -ForegroundColor Gray
        } catch {
            Write-Host "⚠ Tesseract found but not in PATH" -ForegroundColor Yellow
            Write-Host "Adding to PATH..." -ForegroundColor Yellow
            
            # Add to PATH
            $currentPath = [Environment]::GetEnvironmentVariable("Path", "Machine")
            $tesseractDir = "C:\Program Files\Tesseract-OCR"
            
            if ($currentPath -notlike "*$tesseractDir*") {
                $newPath = $currentPath + ";" + $tesseractDir
                [Environment]::SetEnvironmentVariable("Path", $newPath, "Machine")
                Write-Host "✓ Added Tesseract to system PATH" -ForegroundColor Green
                Write-Host "Please restart your terminal to use the updated PATH" -ForegroundColor Cyan
            } else {
                Write-Host "✓ Tesseract already in PATH" -ForegroundColor Green
            }
        }
    } else {
        Write-Host "✗ Tesseract executable not found. Installation may have failed." -ForegroundColor Red
        Write-Host "Expected location: $tesseractPath" -ForegroundColor Gray
    }
    
    # Clean up installer
    Write-Host ""
    Write-Host "Cleaning up..." -ForegroundColor Yellow
    Remove-Item $installerPath -ErrorAction SilentlyContinue
    
    Write-Host ""
    Write-Host "=== Installation Summary ===" -ForegroundColor Cyan
    Write-Host "✓ Tesseract installer downloaded and executed" -ForegroundColor Green
    Write-Host "✓ Installation completed" -ForegroundColor Green
    
    if (Test-Path $tesseractPath) {
        Write-Host "✓ Tesseract executable confirmed" -ForegroundColor Green
    }
    
    Write-Host ""
    Write-Host "Next steps:" -ForegroundColor Cyan
    Write-Host "1. Restart your PowerShell/Command Prompt" -ForegroundColor White
    Write-Host "2. Test with: tesseract --version" -ForegroundColor White
    Write-Host "3. Run your PDF to Excel script: python pdf_to_excel_pymupdf.py" -ForegroundColor White
    
} catch {
    Write-Host ""
    Write-Host "Error during installation: $_" -ForegroundColor Red
    Write-Host ""
    Write-Host "Manual installation steps:" -ForegroundColor Yellow
    Write-Host "1. Download from: $tesseractUrl" -ForegroundColor White
    Write-Host "2. Run the installer as Administrator" -ForegroundColor White
    Write-Host "3. Install to: C:\Program Files\Tesseract-OCR\" -ForegroundColor White
    Write-Host "4. Add to PATH if needed" -ForegroundColor White
}

Write-Host ""
Write-Host "Press any key to exit..." -ForegroundColor Gray
$null = $Host.UI.RawUI.ReadKey("NoEcho,IncludeKeyDown")

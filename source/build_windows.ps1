$ErrorActionPreference = 'Stop'
$root = $PSScriptRoot
Set-Location -LiteralPath $root
$readmeFile = Get-ChildItem -LiteralPath $root -File -Filter 'README*.txt' | Select-Object -First 1
if (-not $readmeFile) { throw 'Package README was not found.' }

$ocrSource = Join-Path $root 'resources\ocr\tesseract'
if (-not (Test-Path -LiteralPath (Join-Path $ocrSource 'tesseract.exe'))) {
    $ocrSource = Join-Path $root '..\ocr\tesseract'
}
if (-not (Test-Path -LiteralPath (Join-Path $ocrSource 'tesseract.exe'))) {
    throw 'Tesseract runtime folder was not found.'
}

py -c "import pymupdf, pytesseract, PIL, openpyxl, PyInstaller"
if ($LASTEXITCODE -ne 0) {
    py -m pip install -r requirements-build.txt
    if ($LASTEXITCODE -ne 0) { throw 'Could not install build dependencies.' }
}

py -m PyInstaller --noconfirm --clean --onedir --windowed --icon assets\pdf_finder.ico --name PeydaPDF --collect-all pymupdf --collect-all openpyxl pdf_keyword_finder.py
if ($LASTEXITCODE -ne 0) { throw 'PyInstaller build failed.' }

$package = Join-Path $root 'dist\PeydaPDF'
$ocrTarget = Join-Path $package 'ocr\tesseract'
New-Item -ItemType Directory -Force -Path $ocrTarget | Out-Null
Copy-Item -Path (Join-Path $ocrSource '*') -Destination $ocrTarget -Recurse -Force
New-Item -ItemType Directory -Force -Path (Join-Path $package 'assets') | Out-Null
New-Item -ItemType Directory -Force -Path (Join-Path $package 'licenses') | Out-Null
New-Item -ItemType Directory -Force -Path (Join-Path $package 'source') | Out-Null
New-Item -ItemType Directory -Force -Path (Join-Path $package 'source\assets') | Out-Null
New-Item -ItemType Directory -Force -Path (Join-Path $package 'source\licenses') | Out-Null
Copy-Item assets\pdf_finder.ico (Join-Path $package 'assets\pdf_finder.ico') -Force
Copy-Item licenses\* (Join-Path $package 'licenses') -Recurse -Force
Copy-Item $readmeFile.FullName (Join-Path $package 'README.txt') -Force
Copy-Item pdf_keyword_finder.py (Join-Path $package 'source\pdf_keyword_finder.py') -Force
Copy-Item requirements.txt (Join-Path $package 'source\requirements.txt') -Force
Copy-Item requirements-build.txt (Join-Path $package 'source\requirements-build.txt') -Force
Copy-Item assets\pdf_finder.ico (Join-Path $package 'source\assets\pdf_finder.ico') -Force
Copy-Item licenses\* (Join-Path $package 'source\licenses') -Recurse -Force
Copy-Item $readmeFile.FullName (Join-Path $package 'source\README.txt') -Force
Copy-Item build_windows.ps1 (Join-Path $package 'source\build_windows.ps1') -Force
Copy-Item build_windows.bat (Join-Path $package 'source\build_windows.bat') -Force

$archive = Join-Path $root 'dist\PeydaPDF-portable.zip'
if (Test-Path -LiteralPath $archive) { Remove-Item -LiteralPath $archive -Force }
Compress-Archive -Path $package -DestinationPath $archive -CompressionLevel Optimal
Write-Host "Portable package created: $archive"

$ErrorActionPreference = "Stop"

$ProjectDirectory = Split-Path -Parent $MyInvocation.MyCommand.Path
Set-Location -LiteralPath $ProjectDirectory

$Python = Join-Path $ProjectDirectory ".venv-build\Scripts\python.exe"
if (-not (Test-Path -LiteralPath $Python)) {
    python -m venv ".venv-build"
}

& $Python -m pip install --upgrade pip
& $Python -m pip install -r "requirements-build.txt"

& $Python -m PyInstaller `
    --noconfirm `
    --clean `
    --onefile `
    --windowed `
    --noupx `
    --name "PDFDataMatrixToCSV" `
    --collect-all "pymupdf" `
    --collect-all "zxingcpp" `
    "main.py"

Write-Host ""
Write-Host "EXE created: $ProjectDirectory\dist\PDFDataMatrixToCSV.exe"

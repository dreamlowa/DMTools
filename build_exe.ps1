$ErrorActionPreference = "Stop"

$ProjectDirectory = Split-Path -Parent $MyInvocation.MyCommand.Path
Set-Location -LiteralPath $ProjectDirectory

# Находим базовый интерпретатор Python (без которого не создать venv).
# Приоритет: явный системный Python в известных местах, затем py.exe, затем python.
function Get-BasePythonCommand {
    $candidates = @(
        "C:\Python313\python.exe",
        "C:\Python312\python.exe",
        "C:\Python311\python.exe",
        "C:\Python310\python.exe"
    )
    foreach ($c in $candidates) {
        if (Test-Path -LiteralPath $c) { return ,@("&", $c) }
    }
    # py.exe — лаунчер версий Python, надёжнее заглушки python из WindowsApps
    $py = Get-Command py.exe -ErrorAction SilentlyContinue
    if ($py) { return ,@("&", "py", "-3") }
    $python = Get-Command python.exe -ErrorAction SilentlyContinue
    if ($python) { return ,@("&", $python.Source) }
    throw "Не найден Python. Установите Python 3.10+ или укажите путь в build_exe.ps1."
}

$Python = Join-Path $ProjectDirectory ".venv-build\Scripts\python.exe"
if (-not (Test-Path -LiteralPath $Python)) {
    $Base = Get-BasePythonCommand
    & $Base[1..($Base.Count - 1)] -m venv ".venv-build"
    if (-not (Test-Path -LiteralPath $Python)) {
        throw "Не удалось создать .venv-build"
    }
}

& $Python -m pip install --upgrade pip
& $Python -m pip install -r "requirements-build.txt"

# Сначала собираем вспомогательный процесс обновления
& $Python -m PyInstaller `
    --noconfirm `
    --clean `
    --onefile `
    --windowed `
    --noupx `
    --name "updater_helper" `
    "updater_helper.py"

# Затем собираем основной EXE, включая helper и zint внутрь
& $Python -m PyInstaller `
    --noconfirm `
    --clean `
    --onefile `
    --windowed `
    --noupx `
    --name "PDFDataMatrixToCSV" `
    --collect-all "pymupdf" `
    --collect-all "zxingcpp" `
    --add-binary "zint.exe;." `
    --add-binary "dist\updater_helper.exe;." `
    --version-file "version_info.txt" `
    "main.py"

Write-Host ""
Write-Host "EXE created: $ProjectDirectory\dist\PDFDataMatrixToCSV.exe"

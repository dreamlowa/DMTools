$ErrorActionPreference = "Stop"

$ProjectDirectory = Split-Path -Parent $MyInvocation.MyCommand.Path
Set-Location -LiteralPath $ProjectDirectory

# РќР°С…РѕРґРёРј Р±Р°Р·РѕРІС‹Р№ РёРЅС‚РµСЂРїСЂРµС‚Р°С‚РѕСЂ Python (Р±РµР· РєРѕС‚РѕСЂРѕРіРѕ РЅРµ СЃРѕР·РґР°С‚СЊ venv).
# РџСЂРёРѕСЂРёС‚РµС‚: СЏРІРЅС‹Р№ СЃРёСЃС‚РµРјРЅС‹Р№ Python РІ РёР·РІРµСЃС‚РЅС‹С… РјРµСЃС‚Р°С…, Р·Р°С‚РµРј py.exe, Р·Р°С‚РµРј python.
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
    # py.exe вЂ” Р»Р°СѓРЅС‡РµСЂ РІРµСЂСЃРёР№ Python, РЅР°РґС‘Р¶РЅРµРµ Р·Р°РіР»СѓС€РєРё python РёР· WindowsApps
    $py = Get-Command py.exe -ErrorAction SilentlyContinue
    if ($py) { return ,@("&", "py", "-3") }
    $python = Get-Command python.exe -ErrorAction SilentlyContinue
    if ($python) { return ,@("&", $python.Source) }
    throw "РќРµ РЅР°Р№РґРµРЅ Python. РЈСЃС‚Р°РЅРѕРІРёС‚Рµ Python 3.10+ РёР»Рё СѓРєР°Р¶РёС‚Рµ РїСѓС‚СЊ РІ build_exe.ps1."
}

$Python = Join-Path $ProjectDirectory ".venv-build\Scripts\python.exe"
if (-not (Test-Path -LiteralPath $Python)) {
    $Base = Get-BasePythonCommand
    & $Base[1..($Base.Count - 1)] -m venv ".venv-build"
    if (-not (Test-Path -LiteralPath $Python)) {
        throw "РќРµ СѓРґР°Р»РѕСЃСЊ СЃРѕР·РґР°С‚СЊ .venv-build"
    }
}

& $Python -m pip install --upgrade pip
& $Python -m pip install -r "requirements-build.txt"

# РЎРЅР°С‡Р°Р»Р° СЃРѕР±РёСЂР°РµРј РІСЃРїРѕРјРѕРіР°С‚РµР»СЊРЅС‹Р№ РїСЂРѕС†РµСЃСЃ РѕР±РЅРѕРІР»РµРЅРёСЏ
& $Python -m PyInstaller `
    --noconfirm `
    --clean `
    --onefile `
    --windowed `
    --noupx `
    --name "updater_helper" `
    "updater_helper.py"

# Р—Р°С‚РµРј СЃРѕР±РёСЂР°РµРј РѕСЃРЅРѕРІРЅРѕР№ EXE, РІРєР»СЋС‡Р°СЏ helper Рё zint РІРЅСѓС‚СЂСЊ
& $Python -m PyInstaller `
    --noconfirm `
    --clean `
    --onefile `
    --windowed `
    --noupx `
    --name "DMTools" `
    --collect-all "pymupdf" `
    --collect-all "zxingcpp" `
    --add-binary "zint.exe;." `
    --add-binary "dist\updater_helper.exe;." `
    --add-data "centrdelfin.ico;." `
    --version-file "version_info.txt" `
    "main.py"

Write-Host ""
Write-Host "EXE created: $ProjectDirectory\dist\DMTools.exe"


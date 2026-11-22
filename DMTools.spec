# -*- mode: python ; coding: utf-8 -*-
from PyInstaller.utils.hooks import collect_all

# tkinter rthook ожидает папки _tcl_data и _tk_data в sys._MEIPASS.
# Python 3.14 хранит библиотеки tcl/tk как zip (libtcl9.0.4.zip/libtk9.0.4.zip);
# они распакованы в ./tcltk и здесь кладутся под ожидаемые rthook имена.
datas = [
    ('centrdelfin.ico', '.'),
    ('tcltk/tcl_library', '_tcl_data'),
    ('tcltk/tk_library', '_tk_data'),
]
binaries = [('zint.exe', '.'), ('dist\\updater_helper.exe', '.')]
hiddenimports = ['scanner', 'scanner_devices', 'scanner_diag']
tmp_ret = collect_all('pymupdf')
datas += tmp_ret[0]; binaries += tmp_ret[1]; hiddenimports += tmp_ret[2]
tmp_ret = collect_all('zxingcpp')
datas += tmp_ret[0]; binaries += tmp_ret[1]; hiddenimports += tmp_ret[2]
for _mod in ('tkinterdnd2', 'openpyxl', 'docx', 'markdown'):
    try:
        tmp_ret = collect_all(_mod)
    except Exception:
        continue
    datas += tmp_ret[0]; binaries += tmp_ret[1]; hiddenimports += tmp_ret[2]


a = Analysis(
    ['main.py'],
    pathex=[],
    binaries=binaries,
    datas=datas,
    hiddenimports=hiddenimports,
    hookspath=[],
    hooksconfig={},
    runtime_hooks=[],
    excludes=[],
    noarchive=False,
    optimize=0,
)
pyz = PYZ(a.pure)

exe = EXE(
    pyz,
    a.scripts,
    a.binaries,
    a.datas,
    [],
    name='DMTools',
    debug=False,
    bootloader_ignore_signals=False,
    strip=False,
    upx=False,
    upx_exclude=[],
    runtime_tmpdir=None,
    console=False,
    disable_windowed_traceback=False,
    argv_emulation=False,
    target_arch=None,
    codesign_identity=None,
    entitlements_file=None,
    version='version_info.txt',
)

"""Р’СЃРїРѕРјРѕРіР°С‚РµР»СЊРЅС‹Р№ РїСЂРѕС†РµСЃСЃ СЃР°РјРѕРѕР±РЅРѕРІР»РµРЅРёСЏ РїСЂРёР»РѕР¶РµРЅРёСЏ.

Р—Р°РїСѓСЃРєР°РµС‚СЃСЏ РѕСЃРЅРѕРІРЅС‹Рј РїСЂРёР»РѕР¶РµРЅРёРµРј (updater.run_self_update) СЃРѕ СЃР»РµРґСѓСЋС‰РёРјРё
Р°СЂРіСѓРјРµРЅС‚Р°РјРё:
    updater_helper.exe <РїСѓС‚СЊ_Рє_С‚РµРєСѓС‰РµРјСѓ_EXE> <URL_РЅРѕРІРѕРіРѕ_EXE>

Р—Р°РґР°С‡Р°: СЃРєР°С‡Р°С‚СЊ РЅРѕРІС‹Р№ EXE РІРѕ РІСЂРµРјРµРЅРЅСѓСЋ РїР°РїРєСѓ, РґРѕР¶РґР°С‚СЊСЃСЏ, РїРѕРєР° РѕСЃРЅРѕРІРЅРѕР№
РїСЂРѕС†РµСЃСЃ Р·Р°РІРµСЂС€РёС‚СЃСЏ, РїРѕРґРјРµРЅРёС‚СЊ С„Р°Р№Р» РїСЂРёР»РѕР¶РµРЅРёСЏ Рё РїРµСЂРµР·Р°РїСѓСЃС‚РёС‚СЊ РµРіРѕ.

Р­С‚Рѕ РѕС‚РґРµР»СЊРЅС‹Р№ РёСЃРїРѕР»РЅСЏРµРјС‹Р№ С„Р°Р№Р», РїРѕСЌС‚РѕРјСѓ РјРѕР¶РµС‚ РїРµСЂРµР·Р°РїРёСЃР°С‚СЊ EXE, РїРѕРєР°
РѕСЃРЅРѕРІРЅРѕРµ РїСЂРёР»РѕР¶РµРЅРёРµ Р·Р°РїСѓС‰РµРЅРѕ.
"""
from __future__ import annotations

import os
import subprocess
import sys
import tempfile
import time
import urllib.request
from pathlib import Path


def download(url: str, dest: Path) -> bool:
    try:
        req = urllib.request.Request(url, headers={"User-Agent": "delfin-etiketki-updater"})
        with urllib.request.urlopen(req, timeout=180) as resp, open(dest, "wb") as out:
            while True:
                chunk = resp.read(65536)
                if not chunk:
                    break
                out.write(chunk)
        return True
    except Exception:
        return False


def wait_for_app_exit(exe_path: Path, timeout: float = 120.0) -> None:
    """Р–РґС‘С‚, РїРѕРєР° РїСЂРѕС†РµСЃСЃ РїСЂРёР»РѕР¶РµРЅРёСЏ Р·Р°РІРµСЂС€РёС‚СЃСЏ. Р¤Р°Р№Р» РјРѕР¶РЅРѕ Р±СѓРґРµС‚ РїРµСЂРµР·Р°РїРёСЃР°С‚СЊ."""
    deadline = time.time() + timeout
    while time.time() < deadline:
        try:
            for proc in subprocess.check_output(
                ["tasklist", "/FI", f"IMAGENAME eq {exe_path.name}", "/FO", "CSV"],
                creationflags=getattr(subprocess, "CREATE_NO_WINDOW", 0),
                text=True,
            ).splitlines()[1:]:
                # РёРјСЏ, pid, СЃРµСЃСЃРёСЏ ...
                if exe_path.name.lower() in proc.lower():
                    break
            else:
                return  # РїСЂРѕС†РµСЃСЃ РЅРµ РЅР°Р№РґРµРЅ вЂ” РјРѕР¶РЅРѕ РѕР±РЅРѕРІР»СЏС‚СЊ
        except Exception:
            pass
        time.sleep(1.0)


def main() -> int:
    if len(sys.argv) < 3:
        return 1
    exe = Path(sys.argv[1])
    url = sys.argv[2]

    tmp = Path(tempfile.gettempdir()) / "delfin_update"
    tmp.mkdir(parents=True, exist_ok=True)
    new_exe = tmp / "DMTools_new.exe"

    if not download(url, new_exe):
        return 2

    wait_for_app_exit(exe)

    # РїРѕРґРјРµРЅСЏРµРј С„Р°Р№Р» РїСЂРёР»РѕР¶РµРЅРёСЏ
    try:
        if exe.exists():
            exe.unlink()
        new_exe.replace(exe)
    except Exception:
        return 3

    # РїРµСЂРµР·Р°РїСѓСЃРєР°РµРј РїСЂРёР»РѕР¶РµРЅРёРµ
    try:
        creationflags = getattr(subprocess, "CREATE_NO_WINDOW", 0)
        subprocess.Popen([str(exe)], creationflags=creationflags, close_fds=True)
    except Exception:
        pass
    return 0


if __name__ == "__main__":
    sys.exit(main())

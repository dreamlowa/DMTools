"""Вспомогательный процесс самообновления приложения.

Запускается основным приложением (updater.run_self_update) со следующими
аргументами:
    updater_helper.exe <путь_к_текущему_EXE> <URL_нового_EXE>

Задача: скачать новый EXE во временную папку, дождаться, пока основной
процесс завершится, подменить файл приложения и перезапустить его.

Это отдельный исполняемый файл, поэтому может перезаписать EXE, пока
основное приложение запущено.
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


def wait_for_app_exit(exe_path: Path, pid_hint: str, timeout: float = 120.0) -> None:
    """Ждёт, пока процесс приложения завершится. Файл можно будет перезаписать."""
    deadline = time.time() + timeout
    while time.time() < deadline:
        try:
            for proc in subprocess.check_output(
                ["tasklist", "/FI", f"IMAGENAME eq {exe_path.name}", "/FO", "CSV"],
                creationflags=getattr(subprocess, "CREATE_NO_WINDOW", 0),
                text=True,
            ).splitlines()[1:]:
                # имя, pid, сессия ...
                if exe_path.name.lower() in proc.lower():
                    break
            else:
                return  # процесс не найден — можно обновлять
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
    new_exe = tmp / "PDFDataMatrixToCSV_new.exe"

    if not download(url, new_exe):
        return 2

    wait_for_app_exit(exe)

    # подменяем файл приложения
    try:
        if exe.exists():
            exe.unlink()
        new_exe.replace(exe)
    except Exception:
        return 3

    # перезапускаем приложение
    try:
        creationflags = getattr(subprocess, "CREATE_NO_WINDOW", 0)
        subprocess.Popen([str(exe)], creationflags=creationflags, close_fds=True)
    except Exception:
        pass
    return 0


if __name__ == "__main__":
    sys.exit(main())
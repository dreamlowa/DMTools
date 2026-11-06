"""Самообновление приложения.

Механизм: программа обращается к манифесту обновлений на GitHub
(адрес задаётся в константах). Эталонный манифест выглядит так:

    {
      "version": "1.5.0",
      "url": "https://github.com/USER/REPO/releases/download/v1.5.0/DMTools.exe",
      "notes": "Краткое описание изменений"
    }

Если версия из манифеста выше текущей (APP_VERSION), программа предлагает
обновление. Обновление выполняется вспомогательным процессом (helper),
который скачивает новый EXE, ждёт закрытия основного приложения,
подменяет файл и перезапускает приложение.
"""
from __future__ import annotations

import json
import os
import subprocess
import sys
import tempfile
import urllib.request
from pathlib import Path

from version import APP_VERSION

# Адрес манифеста обновлений на GitHub. Замените на фактический URL.
UPDATE_MANIFEST_URL = (
    "https://raw.githubusercontent.com/dreamlowa/delfin-etiketki/main/update.json"
)
# Имя вспомогательного помощника обновления (собирается отдельно в EXE)
UPDATER_HELPER = "updater_helper.exe"


def parse_version(v: str) -> tuple[int, int, int]:
    """Разбирает строку версии 'X.Y.Z' в кортеж целых чисел."""
    parts = (v or "").strip().split(".")
    nums: list[int] = []
    for p in parts[:3]:
        try:
            nums.append(int(p))
        except ValueError:
            nums.append(0)
    while len(nums) < 3:
        nums.append(0)
    return (nums[0], nums[1], nums[2])


def is_newer(latest: str, current: str) -> bool:
    """True, если latest строго новее current."""
    return parse_version(latest) > parse_version(current)


def fetch_manifest(timeout: int = 15) -> dict | None:
    """Скачивает манифест обновлений с GitHub. Возвращает dict или None."""
    try:
        req = urllib.request.Request(
            UPDATE_MANIFEST_URL, headers={"User-Agent": "delfin-etiketki-updater"}
        )
        with urllib.request.urlopen(req, timeout=timeout) as resp:
            data = json.loads(resp.read().decode("utf-8"))
        if "version" in data and "url" in data:
            return data
        return None
    except Exception:
        return None


def check_for_update(manifest: dict | None = None) -> dict:
    """Проверяет, доступно ли обновление.

    Возвращает dict вида:
      {"status": "ok", "update": {...} | None, "current": "1.8.2",
       "latest": "1.8.2" | None}
    либо {"status": "error", "message": "..."} при сбое сети/разбора.

    Ключ "update" содержит манифест, только если найдена более новая версия.
    """
    m = manifest if manifest is not None else fetch_manifest()
    if not m:
        return {
            "status": "error",
            "message": "Не удалось получить данные об обновлениях. "
                       "Проверьте подключение к интернету и попробуйте снова.",
        }
    latest = m.get("version", "")
    current = APP_VERSION
    if is_newer(latest, current):
        return {"status": "ok", "update": m, "current": current, "latest": latest}
    return {"status": "ok", "update": None, "current": current, "latest": latest}


def app_exe_path() -> Path:
    """Путь к главному исполняемому файлу приложения."""
    return Path(sys.executable).resolve()


def helper_exe_path() -> Path | None:
    """Путь к вспомогательному обновлятору (рядом с EXE или в _MEIPASS)."""
    base = Path(getattr(sys, "_MEIPASS", Path(sys.executable).resolve().parent))
    cands = [
        base / UPDATER_HELPER,
        Path(sys.executable).resolve().parent / UPDATER_HELPER,
    ]
    for c in cands:
        if c.is_file():
            return c
    return None


def download_file(url: str, dest: Path, timeout: int = 120) -> bool:
    """Скачивает файл по URL в dest. Возвращает True при успехе."""
    try:
        req = urllib.request.Request(url, headers={"User-Agent": "delfin-etiketki-updater"})
        with urllib.request.urlopen(req, timeout=timeout) as resp, open(
            dest, "wb"
        ) as out:
            while True:
                chunk = resp.read(65536)
                if not chunk:
                    break
                out.write(chunk)
        return True
    except Exception:
        return False


def run_self_update(url: str) -> str:
    """Запускает вспомогательный процесс самообновления.

    Возвращает пустую строку при успешном запуске helper-а, либо текст ошибки."""
    helper = helper_exe_path()
    if helper is None:
        return f"Не найден {UPDATER_HELPER} рядом с программой."
    exe = app_exe_path()
    try:
        # Сообщаем helper-у: путь к файлу, который нужно обновить,
        # и URL, откуда качать новую версию.
        cmd = [str(helper), str(exe), url]
        creationflags = getattr(subprocess, "CREATE_NO_WINDOW", 0)
        subprocess.Popen(
            cmd,
            creationflags=creationflags,
            close_fds=True,
        )
        return ""
    except Exception as exc:
        return f"Не удалось запустить обновление: {exc}"
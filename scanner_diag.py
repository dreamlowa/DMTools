"""Диагностика подключения сканера (Windows, без внешних зависимостей).

Собирает в текстовый журнал полный отчёт о том, что Windows видит как
устройства ввода и последовательные порты — чтобы разобраться, почему
сканер не определяется или не присылает данные.

Всё реализовано через встроенные средства Windows (реестр + WMI через
PowerShell) и стандартную библиотеку Python, чтобы не раздувать EXE.
"""

from __future__ import annotations

import subprocess
import sys
import winreg
from datetime import datetime


def _run_ps(code: str) -> str:
    """Выполняет PowerShell-код и возвращает stdout (UTF-8)."""
    ps = ("[Console]::OutputEncoding=[System.Text.Encoding]::UTF8; " + code)
    args = ["powershell", "-NoProfile", "-NonInteractive", "-Command", ps]
    try:
        res = subprocess.run(args, capture_output=True,
                             timeout=20, creationflags=0x08000000)
        return res.stdout.decode("utf-8", errors="replace").strip()
    except Exception as exc:
        return f"(powershell ошибка: {exc})"


def _cim_table(where: str, props: list[str]) -> str:
    """Win32_PnPEntity в виде таблицы Name | PNPClass | DeviceID."""
    cols = ", ".join(props)
    code = (
        f"Get-CimInstance Win32_PnPEntity -ErrorAction SilentlyContinue "
        f"| Where-Object {{ {where} }} "
        f"| Select-Object {cols} "
        f"| Format-Table -AutoSize | Out-String -Width 4096"
    )
    return _run_ps(code)


def collect_diagnostic() -> str:
    """Собирает полный журнал диагностики сканера и возвращает строку."""
    lines: list[str] = []
    lines.append("===== ЖУРНАЛ ДИАГНОСТИКИ СКАНЕРА =====")
    lines.append(f"Собран: {datetime.now().strftime('%Y-%m-%d %H:%M:%S')}")
    lines.append(f"Python: {sys.version.split()[0]}")
    lines.append("")

    # --- COM-порты из реестра ---
    lines.append("--- 1. COM-порты (реестр) ---")
    coms = []
    try:
        key = winreg.OpenKey(winreg.HKEY_LOCAL_MACHINE,
                             r"HARDWARE\DEVICEMAP\SERIALCOMM")
    except OSError:
        coms = "(не доступны)"
    else:
        try:
            for i in range(winreg.QueryInfoKey(key)[1]):
                _n, value, _t = winreg.EnumValue(key, i)
                if value:
                    coms.append(str(value))
        finally:
            winreg.CloseKey(key)
    lines.append(", ".join(coms) if coms else "COM-портов не найдено")
    lines.append("")

    # --- Все устройства ввода с VID/PID ---
    lines.append("--- 2. Устройства ввода и HID с VID:PID ---")
    lines.append(_cim_table(
        "$_.DeviceID -match 'VID_'",
        ["Name", "PNPClass", "DeviceID"]))
    lines.append("")

    # --- Клавиатуры ---
    lines.append("--- 3. Клавиатуры (Win32_Keyboard) ---")
    lines.append(_run_ps(
        "Get-CimInstance Win32_Keyboard -ErrorAction SilentlyContinue "
        "| Select-Object Name, DeviceID "
        "| Format-Table -AutoSize | Out-String -Width 4096"))
    lines.append("")

    # --- Все HID-устройства (полный список, даже без VID) ---
    lines.append("--- 4. Все HID-устройства ---")
    lines.append(_cim_table("$_.PNPClass -eq 'HIDClass'",
                            ["Name", "DeviceID"]))
    lines.append("")

    # --- USB-устройства ---
    lines.append("--- 5. USB-устройства ---")
    lines.append(_cim_table(
        "$_.DeviceID -like 'USB\\\\*' -and $_.Name -notmatch 'концентратор|root hub|Root Hub'",
        ["Name", "PNPClass", "DeviceID"]))
    lines.append("")

    # --- Композитные USB и их статус ---
    lines.append("--- 6. Статус устройств ввода ---")
    lines.append(_run_ps(
        "Get-CimInstance Win32_PnPEntity -ErrorAction SilentlyContinue "
        "| Where-Object { $_.Name -match 'ввод|scanner|сканер|barcode|штрих|HID|Композит' "
        "} | Select-Object Name, PNPClass, Status "
        "| Format-Table -AutoSize | Out-String -Width 4096"))
    lines.append("")
    lines.append("===== КОНЕЦ ЖУРНАЛА =====")
    return "\n".join(lines)


if __name__ == "__main__":
    print(collect_diagnostic())
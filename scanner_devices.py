"""Обнаружение сканеров штрихкодов как оборудования (Windows).

Сканер может быть подключён двумя способами:
  * Клавиатура (USB HID Keyboard) — определяется Windows как устройство
    ввода, COM-порта нет, код приходит как обычный ввод + Enter.
  * COM/Serial (в т.ч. Virtual COM через USB-кабель) — у сканера есть
    номер COM-порта; программа должна сама открыть порт и читать данные.

Модуль перечисляет доступное оборудование без внешних зависимостей
(pyserial / wmi не требуются):
  * COM-порты читаются из реестра HKLM\\HARDWARE\\DEVICEMAP\\SERIALCOMM;
  * читаемое имя устройства и VID:PID берутся из Win32_PnPEntity
    (через вызов Get-CimInstance в подпроцессе, т.к. wmi-библиотеки нет).

Это discovery-функции: они ничего не читают со сканера, а лишь показывают,
что Windows видит как последовательные устройства и USB-устройства ввода.
"""

from __future__ import annotations

import re
import subprocess
import winreg
from dataclasses import dataclass, field


@dataclass
class PortInfo:
    """Информация об одном последовательном (COM) порте."""
    device: str            # например "COM3"
    name: str = ""         # человекочитаемое имя устройства
    vid: str = ""          # "0A5F" без "VID_"
    pid: str = ""          # "1234" без "PID_"
    description: str = ""  # полное описание из PnP


@dataclass
class HidInfo:
    """USB/HID-устройство ввода (может быть сканером в клавиатурном режиме)."""
    name: str        # человекочитаемое имя
    vid: str = ""
    pid: str = ""
    instance: str = field(default="")


# Производители, чьи сканеры штрихкодов чаще всего встречаются.
# Список не исчерпывающий — используется только для подсказки в UI.
KNOWN_SCANNER_VIDS = {
    "0536",  # Hand Held Products / Honeywell
    "05e0",  # Symbol / Zebra
    "0a5f",  # Zebra (Motorola)
    "0463",  # Datalogic
    "05c6",  # Qualcomm (некоторые встраиваемые)
    "0602",  # Mindeo
}
# Из объединённого списка исключаем то, что явно не сканер
EXCLUDED_WORDS = ("Realtek", "Signetics", "концентратор", "hub", "хаб",
                  "Wireless Controller", "GamePad", "Джойстик")


def _run_cim(query: str, props: list[str]) -> list[dict]:
    """Выполняет запрос Win32_PnPEntity и возвращает список словарей.
    Работает через subprocess + Get-CimInstance (wmi-библиотеки не требуется).
    Возвращает [] при любой ошибке (команды нет, недоступен и т.п.).
    """
    ps = (
        "[Console]::OutputEncoding=[System.Text.Encoding]::UTF8; "
        f"Get-CimInstance Win32_PnPEntity -ErrorAction SilentlyContinue "
        f"| Where-Object {{ {query} }} "
        f"| Select-Object {', '.join(props)} "
        f"| ConvertTo-Json -Compress"
    )
    args = ["powershell", "-NoProfile", "-NonInteractive", "-Command", ps]
    try:
        # PowerShell переведён на UTF-8 вывод, поэтому декодируем как utf-8.
        res = subprocess.run(args, capture_output=True,
                             timeout=15, creationflags=0x08000000)
        out = res.stdout.decode("utf-8", errors="replace")
    except Exception:
        return []
    if not out or not out.strip():
        return []
    try:
        import json
        data = json.loads(out)
    except Exception:
        return []
    if isinstance(data, dict):
        data = [data]
    return data if isinstance(data, list) else []


def list_com_ports() -> list[PortInfo]:
    """Возвращает список COM-портов, зарегистрированных в системе.

    Берёт имена портов из реестра, затем для каждого ищет читаемое имя
    устройства и VID/PID в Win32_PnPEntity.
    """
    ports: list[PortInfo] = []
    try:
        key = winreg.OpenKey(winreg.HKEY_LOCAL_MACHINE,
                             r"HARDWARE\DEVICEMAP\SERIALCOMM")
    except OSError:
        return ports
    try:
        count = winreg.QueryInfoKey(key)[1]
        for i in range(count):
            try:
                _name, value, _type = winreg.EnumValue(key, i)
            except OSError:
                continue
            if value and str(value).startswith("COM"):
                ports.append(PortInfo(device=str(value)))
    finally:
        winreg.CloseKey(key)

    if not ports:
        return ports
    ports.sort(key=lambda p: _com_sort_key(p.device))

    # имена устройств и VID/PID из PnP
    pnp = _run_cim(
        "($_.PNPClass -in @('Ports','Modem','USB') -or $_.DeviceID -like 'USB\\\\*')",
        ["Name", "DeviceID", "PNPClass"])
    for p in ports:
        for item in pnp:
            devid = item.get("DeviceID") or ""
            name = item.get("Name") or ""
            # ищем запись, чей DeviceID упоминает наш порт либо имя содержит "COMx"
            if f"({p.device.upper()})" in name.upper() or f"\\{p.device.upper()}" in devid.upper():
                p.name = name
                _extract_vidpid(devid, p)
                p.description = name
                break
        if not p.name:
            p.name = p.device
    return ports


def list_hid_keyboards() -> list[HidInfo]:
    """Возвращает USB/HID-устройства ввода (возможные сканеры в режиме
    клавиатуры). Фильтрует по PNPClass Keyboard/HIDClass и наличию VID/PID."""
    pnp = _run_cim(
        "($_.PNPClass -in @('Keyboard','HIDClass','USB') -or "
        "$_.DeviceID -like 'USB\\\\*')",
        ["Name", "DeviceID", "PNPClass"])
    result: list[HidInfo] = []
    seen: set[str] = set()
    for item in pnp:
        devid = item.get("DeviceID") or ""
        name = item.get("Name") or ""
        vid, pid = _parse_vidpid(devid)
        if not vid:
            continue  # без VID/PID не можем идентифицировать устройство
        if any(w in name for w in EXCLUDED_WORDS):
            continue
        key = f"{vid}:{pid}"
        if key in seen:
            continue
        seen.add(key)
        result.append(HidInfo(name=name or f"USB устройство {vid}:{pid}",
                              vid=vid, pid=pid, instance=devid))
    result.sort(key=lambda h: h.name.lower())
    return result


def is_likely_scanner_vid(vid: str) -> bool:
    """True, если производитель VID известен как производитель сканеров."""
    return vid.lower() in KNOWN_SCANNER_VIDS


def _com_sort_key(device: str) -> int:
    m = re.search(r"(\d+)$", device)
    return int(m.group(1)) if m else 0


def _extract_vidpid(devid: str, port: PortInfo) -> None:
    vid, pid = _parse_vidpid(devid)
    port.vid, port.pid = vid, pid


def _parse_vidpid(devid: str) -> tuple[str, str]:
    r"""Возвращает (VID, PID) из строки вида 'USB\VID_0A5F&PID_1234'."""
    m = re.search(r"VID[_&]([0-9A-Fa-f]{4})", devid)
    v = re.search(r"PID[_&]([0-9A-Fa-f]{4})", devid)
    vid = m.group(1).lower() if m else ""
    pid = v.group(1).lower() if v else ""
    return vid, pid


if __name__ == "__main__":
    print("COM-порты:")
    for p in list_com_ports():
        print(f"  {p.device} | {p.name} | VID:{p.vid} PID:{p.pid}")
    print("HID/USB устройства ввода:")
    for h in list_hid_keyboards():
        mark = "  <-- возможно сканер" if is_likely_scanner_vid(h.vid) else ""
        print(f"  {h.name} | VID:{h.vid} PID:{h.pid}{mark}")
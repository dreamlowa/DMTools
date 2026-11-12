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
# Используется, чтобы предлагать устройства в настройках.
KNOWN_SCANNER_VIDS = {
    "0536",  # Hand Held Products / Honeywell
    "05e0",  # Symbol / Zebra
    "0a5f",  # Zebra (Motorola)
    "0463",  # Datalogic
    "0602",  # Mindeo
    "2600",  # ZIH Corp / Zebra HD
    "067b",  # Prolific (часть USB-serial адаптеров сканеров)
    "10c4",  # Silicon Labs CP210x (USB-serial, часто у сканеров)
}

# Слова, по которым устройство НЕ похоже на сканер (фильтр по имени).
# Служат, чтобы не показывать мыши, клавиатуры, тачпады, наушники,
# камеры, принтеры, концентраторы и т.п.
NON_SCANNER_WORDS = (
    "мышь", "mouse", "клавиатур", "keyboard", "тачпад", "touchpad",
    "сенсорная панель", "наушник", "headphone", "гарнитур", "headset",
    "камер", "camera", "микрофон", "microphone", "динамик", "speaker",
    "принтер", "printer", "сканер изображ", "image scanner", "webcam",
    "концентратор", "hub", "хаб", "ридер", "card reader", "кардридер",
    "Realtek", "жёсткий диск", "диск", "hard disk", "флеш", "flash",
    "телефон", "phone", "контроллер", "controller", "джойстик", "joystick",
    "gaming", "LED", "подсветк", "RGB", "монитор", "monitor", "bt",
    "bluetooth", "радио", "rf", "беспроводн", "шнур", "кабель", "cable",
    "заряд", "charg", "usb-концентратор",
)

# VID, которые почти наверняка НЕ сканеры (бытовая периферия, встроенное ж/д)
NON_SCANNER_VIDS = {
    "0bda",  # Realtek — кардридеры/аудио/USB-радио
    "1ea7",  # дешёвая универсальная USB-периферия (мыши/клавиатуры)
    "3c4d",  # встроенная клавиатура/тачпад ноутбуков
    "0d62",  # Darfon (клавиатуры)
    "046d",  # Logitech (мыши/клавиатуры/камеры)
    "093a",  # Pixart (мыши/веб-камеры)
    "1235",  # USB-пульты
}


def _is_scanner_like(name: str, vid: str) -> bool:
    """Решает, похоже ли устройство на сканер штрихкодов.

    Показывает TRUE (т.е. это кандидат на сканер), если:
      * имя упоминает сканер/штрих-код, ИЛИ
      * производитель VID входит в список производителей сканеров,
    и при этом устройство НЕ выглядит как мышь/клавиатура/тачпад/
    наушники/принтер/камера/хаб и т.п.
    """
    low = name.lower()
    # явно не сканер по имени
    if any(w in low for w in NON_SCANNER_WORDS):
        return False
    # известный не-сканер по VID
    if vid in NON_SCANNER_VIDS:
        return False
    # по имени похоже на сканер
    if any(k in low for k in ("сканер", "штрих", "barcode", "scanner",
                              "маркировк", "datamatrix", "data matrix")):
        return True
    # производитель известен как производитель сканеров
    if vid in KNOWN_SCANNER_VIDS:
        return True
    return False


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
                vid, pid = _parse_vidpid(devid)
                # COM-порт показываем, если он похож на сканер ИЛИ о нём
                # ничего не известно (нет имени/VID) — тогда это просто порт
                if not _is_scanner_like(name, vid) and vid:
                    break
                p.name = name
                p.vid, p.pid = vid, pid
                p.description = name
                break
        if not p.name:
            p.name = p.device
    return ports


def list_hid_keyboards() -> list[HidInfo]:
    """Возвращает USB/HID-устройства, похожие на сканер (клавиатурный режим).

    Показывает только те устройства, которые могут быть сканером штрихкодов:
    мыши, клавиатуры, тачпады, наушники, камеры, принтеры, концентраторы и
    прочая бытовая периферия отсекаются через _is_scanner_like.
    """
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
        if not _is_scanner_like(name, vid):
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
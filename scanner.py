"""Чтение кодов со сканера через COM-порт (Windows, без внешних зависимостей).

Некоторые сканеры работают не как клавиатура, а как последовательное
устройство (COM/Serial, в т.ч. Virtual COM через USB-кабель). Тогда
программа должна сама открыть порт, читать байты и собирать из них код.

Порт открывается напрямую через WinAPI (CreateFile/ReadFile) средствами
ctypes — pyserial не требуется. Данные читаются в отдельном потоке до
символа-суффикса (CR / LF / CRLF), после чего строка считается кодом.

	Чтение ведётся по одному байту за раз, что медленно, но порт сканера
	низкоскоростной и для DataMatrix (~100-200 символов) это приемлемо.
"""

from __future__ import annotations

import ctypes
import ctypes.wintypes as wt
import threading

# --- константы WinAPI ---
GENERIC_READ = 0x80000000
GENERIC_WRITE = 0x40000000
OPEN_EXISTING = 3
FILE_ATTRIBUTE_NORMAL = 0x80
CREATE_NO_WINDOW = 0x08000000

ERROR_ACCESS_DENIED = 5
ERROR_FILE_NOT_FOUND = 2

INVALID_HANDLE_VALUE = ctypes.c_void_p(-1).value


class _OVERLAPPED(ctypes.Structure):
    _fields_ = [
        ("Internal", wt.ULONG),
        ("InternalHigh", wt.ULONG),
        ("Offset", wt.DWORD),
        ("OffsetHigh", wt.DWORD),
        ("hEvent", wt.HANDLE),
    ]


_kernel32 = ctypes.WinDLL("kernel32", use_last_error=True)

_kernel32.CreateFileW.argtypes = [wt.LPCWSTR, wt.DWORD, wt.DWORD, ctypes.c_void_p,
                                  wt.DWORD, wt.DWORD, wt.HANDLE]
_kernel32.CreateFileW.restype = wt.HANDLE
_kernel32.ReadFile.restype = wt.BOOL
_kernel32.SetCommTimeouts.restype = wt.BOOL
_kernel32.CloseHandle.restype = wt.BOOL
_kernel32.SetCommState.restype = wt.BOOL
_kernel32.GetCommState.restype = wt.BOOL


class COMMTIMEOUTS(ctypes.Structure):
    _fields_ = [
        ("ReadIntervalTimeout", wt.DWORD),
        ("ReadTotalTimeoutMultiplier", wt.DWORD),
        ("ReadTotalTimeoutConstant", wt.DWORD),
        ("WriteTotalTimeoutMultiplier", wt.DWORD),
        ("WriteTotalTimeoutConstant", wt.DWORD),
    ]


class DCB(ctypes.Structure):
    _fields_ = [
        ("DCBlength", wt.DWORD),
        ("BaudRate", wt.DWORD),
        ("fBinary", wt.DWORD),
        ("fParity", wt.DWORD),
        ("fOutxCtsFlow", wt.DWORD),
        ("fOutxDsrFlow", wt.DWORD),
        ("fDtrControl", wt.DWORD),
        ("fDsrSensitivity", wt.DWORD),
        ("fTXContinueOnXoff", wt.DWORD),
        ("fOutX", wt.DWORD),
        ("fInX", wt.DWORD),
        ("fErrorChar", wt.DWORD),
        ("fNull", wt.DWORD),
        ("fRtsControl", wt.DWORD),
        ("fAbortOnError", wt.DWORD),
        ("fDummy2", wt.DWORD),
        ("wReserved", wt.WORD),
        ("XonLim", wt.WORD),
        ("XoffLim", wt.WORD),
        ("ByteSize", wt.BYTE),
        ("Parity", wt.BYTE),
        ("StopBits", wt.BYTE),
        ("XonChar", ctypes.c_char),
        ("XoffChar", ctypes.c_char),
        ("ErrorChar", ctypes.c_char),
        ("EofChar", ctypes.c_char),
        ("EvtChar", ctypes.c_char),
        ("wReserved1", wt.WORD),
    ]


class SerialScanner:
    """Открывает COM-порт сканера и читает коды в фоновом потоке.

    При каждом завершённом коде вызывает on_code(code). При ошибке порта —
    on_error(message). close() останавливает чтение и закрывает порт.
    """

    # базовая скорость: символы ASCII 8N1. Для GTIN хватает 9600.
    DEFAULT_BAUD = 9600

    def __init__(self, port: str, baud: int = DEFAULT_BAUD,
                 on_code=None, on_error=None):
        self.port = port
        self.baud = baud
        self.on_code = on_code
        self.on_error = on_error
        self._handle = None
        self._stop = threading.Event()
        self._thread: threading.Thread | None = None
        self._lock = threading.Lock()

    # ------------------------------------------------------------------ #
    # открытие/закрытие
    # ------------------------------------------------------------------ #
    def open(self) -> str | None:
        """Открывает порт Для чтения. Возвращает None при успехе, иначе
        текст ошибки (читаемый)."""
        portw = "\\\\.\\" + self.port  # например \\.\COM3
        handle = _kernel32.CreateFileW(
            portw, GENERIC_READ | GENERIC_WRITE, 0, None,
            OPEN_EXISTING, FILE_ATTRIBUTE_NORMAL | CREATE_NO_WINDOW, None)
        if handle in (INVALID_HANDLE_VALUE, None):
            err = ctypes.get_last_error()
            if err == ERROR_FILE_NOT_FOUND:
                return f"Порт {self.port} не найден. Проверьте подключение сканера."
            if err == ERROR_ACCESS_DENIED:
                return f"Порт {self.port} занят другой программой."
            return f"Не удалось открыть порт {self.port} (код {err})."
        self._handle = handle
        self._configure()
        # приём сразу: чтение запускается в open()
        self._stop.clear()
        self._thread = threading.Thread(target=self._reader, daemon=True)
        self._thread.start()
        return None

    def _configure(self) -> None:
        try:
            dcb = DCB()
            dcb.DCBlength = ctypes.sizeof(DCB)
            # читаем текущие параметры и подстраиваем скорость
            _kernel32.GetCommState(self._handle, ctypes.byref(dcb))
            dcb.BaudRate = self.baud
            dcb.ByteSize = 8
            dcb.Parity = 0
            dcb.StopBits = 0  # ones
            dcb.fBinary = 1
            dcb.fParity = 0
            dcb.fNull = 0
            dcb.fRtsControl = 0
            dcb.fDtrControl = 0
            _kernel32.SetCommState(self._handle, ctypes.byref(dcb))
        except Exception:
            pass  # если не удалось настроить — оставляем значения по умолчанию
        try:
            to = COMMTIMEOUTS()
            to.ReadIntervalTimeout = 0
            to.ReadTotalTimeoutMultiplier = 0
            to.ReadTotalTimeoutConstant = 100  # 100 мс
            to.WriteTotalTimeoutMultiplier = 0
            to.WriteTotalTimeoutConstant = 0
            _kernel32.SetCommTimeouts(self._handle, ctypes.byref(to))
        except Exception:
            pass

    def close(self) -> None:
        """Останавливает поток чтения и закрывает порт (идемпотентно)."""
        self._stop.set()
        thread = self._thread
        if thread is not None and thread is not threading.current_thread():
            thread.join(timeout=1.5)
        with self._lock:
            h = self._handle
            self._handle = None
            if h not in (None, INVALID_HANDLE_VALUE):
                try:
                    _kernel32.CloseHandle(h)
                except Exception:
                    pass

    # ------------------------------------------------------------------ #
    # чтение
    # ------------------------------------------------------------------ #
    def _reader(self) -> None:
        buf = bytearray()
        while not self._stop.is_set():
            byte = self._read_byte()
            if byte is None:
                if self._stop.is_set():
                    break
                continue
            buf.append(byte)
            # суффикс конца строки: CR (13) или CRLF (13,10)
            if byte in (13, 10):
                line = bytes(buf).strip(b"\r\n")
                buf.clear()
                if line:
                    try:
                        code = line.decode("utf-8", errors="replace")
                    except Exception:
                        code = ""
                    if code and self.on_code:
                        self.on_code(code)
        self.close()

    def _read_byte(self) -> int | None:
        if self._handle in (None, INVALID_HANDLE_VALUE):
            return None
        ch = ctypes.c_byte()
        nread = wt.DWORD(0)
        ok = _kernel32.ReadFile(self._handle, ctypes.byref(ch), 1,
                                ctypes.byref(nread), None)
        if not ok:
            err = ctypes.get_last_error()
            # порт закрыт снаружи — просто выходим
            if err == ERROR_ACCESS_DENIED:
                if self.on_error:
                    self.on_error(f"Порт {self.port} закрыт.")
                return None
            return None
        if nread.value == 0:
            return None
        return int(ch.value) & 0xFF


def open_and_read(port: str, on_code, on_error,
                  baud: int = SerialScanner.DEFAULT_BAUD) -> SerialScanner:
    """Удобная обёртка: создаёт SerialScanner, открывает порт.

    Возвращает объект. Проверяйте scanner.last_error: если порт не открылся,
    объект пустой и его можно проигнорировать."""
    scanner = SerialScanner(port, baud=baud, on_code=on_code, on_error=on_error)
    err = scanner.open()
    if err:
        if on_error:
            on_error(err)
        scanner.close()
        return scanner
    return scanner


if __name__ == "__main__":
    import sys
    port = sys.argv[1] if len(sys.argv) > 1 else "COM3"
    print(f"Слушаю порт {port}. Наведите сканер. Ctrl+C для выхода.")

    def _code(c: str) -> None:
        print("КОД:", repr(c))

    def _err(e: str) -> None:
        print("ОШИБКА:", e)

    sc = open_and_read(port, _code, _err)
    if not sc or sc._handle is None:
        sys.exit(1)
    try:
        while True:
            import time
            time.sleep(1)
    except KeyboardInterrupt:
        sc.close()
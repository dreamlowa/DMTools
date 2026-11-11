from __future__ import annotations

import csv
import os
import string
import sys
import tempfile
import threading
from dataclasses import dataclass
from pathlib import Path
from typing import Callable

import pymupdf
from PIL import Image
import zxingcpp


GS = b"\x1d"
DEFAULT_DPIS = (216, 288, 360, 400)
GS1_X_CHARACTERS = frozenset(
    string.ascii_letters + string.digits + "!\"%&'()*+,-./:;<=>?_"
)

LogCallback = Callable[[str], None]
PageProgressCallback = Callable[[int, int], None]


class ConversionError(RuntimeError):
    pass


class ConversionCancelled(RuntimeError):
    pass


@dataclass(frozen=True)
class MarkingCode:
    page: int
    gtin: str
    serial: str
    crypto_key_91: str
    crypto_code_92: str
    raw: bytes


@dataclass(frozen=True)
class ConversionResult:
    pdf_path: Path
    csv_path: Path
    page_count: int
    code_count: int
    unique_code_count: int
    gtins: dict[str, int]


@dataclass(frozen=True)
class PdfDecodeResult:
    pdf_path: Path
    page_count: int
    codes: tuple[MarkingCode, ...]

    @property
    def code_count(self) -> int:
        return len(self.codes)

    @property
    def unique_code_count(self) -> int:
        return len({code.raw for code in self.codes})

    @property
    def gtins(self) -> dict[str, int]:
        counts: dict[str, int] = {}
        for code in self.codes:
            counts[code.gtin] = counts.get(code.gtin, 0) + 1
        return counts


def gtin_check_digit_is_valid(gtin: str) -> bool:
    if len(gtin) != 14 or not gtin.isdigit():
        return False
    body = [int(char) for char in gtin[:-1]]
    weighted_sum = sum(
        digit * (3 if (len(body) - index) % 2 else 1)
        for index, digit in enumerate(body)
    )
    return (-weighted_sum) % 10 == int(gtin[-1])


def _validate_gs1_text(ai: str, value: str, minimum: int, maximum: int) -> None:
    if not minimum <= len(value) <= maximum:
        raise ValueError(
            f"AI {ai} length must be {minimum}..{maximum}, got {len(value)}"
        )
    invalid = sorted(set(value) - GS1_X_CHARACTERS)
    if invalid:
        rendered = ", ".join(repr(char) for char in invalid)
        raise ValueError(f"AI {ai} contains invalid GS1 characters: {rendered}")


def parse_honest_sign_code(raw: bytes, page_number: int) -> MarkingCode:
    # Honest Sign layout:
    # 01 + GTIN(14) + 21 + serial + GS + 91 + key + GS + 92 + signature
    if not raw.startswith(b"01") or len(raw) < 18:
        raise ValueError("AI 01 with a 14-digit GTIN was not found")

    gtin_bytes = raw[2:16]
    if raw[16:18] != b"21":
        raise ValueError("AI 21 was not found after AI 01")

    parts = raw[18:].split(GS)
    if len(parts) != 3:
        raise ValueError(f"expected two GS separators, found {len(parts) - 1}")

    serial_bytes, ai91_part, ai92_part = parts
    if not ai91_part.startswith(b"91"):
        raise ValueError("AI 91 was not found after the first GS separator")
    if not ai92_part.startswith(b"92"):
        raise ValueError("AI 92 was not found after the second GS separator")

    try:
        gtin = gtin_bytes.decode("ascii")
        serial = serial_bytes.decode("ascii")
        crypto_key_91 = ai91_part[2:].decode("ascii")
        crypto_code_92 = ai92_part[2:].decode("ascii")
    except UnicodeDecodeError as exc:
        raise ValueError("decoded code contains non-ASCII bytes") from exc

    if not gtin_check_digit_is_valid(gtin):
        raise ValueError(f"invalid GTIN check digit: {gtin}")
    _validate_gs1_text("21", serial, 1, 20)
    _validate_gs1_text("91", crypto_key_91, 1, 90)
    _validate_gs1_text("92", crypto_code_92, 1, 90)

    return MarkingCode(
        page=page_number,
        gtin=gtin,
        serial=serial,
        crypto_key_91=crypto_key_91,
        crypto_code_92=crypto_code_92,
        raw=raw,
    )


def render_page(page: pymupdf.Page, dpi: int) -> Image.Image:
    scale = dpi / 72
    pixmap = page.get_pixmap(
        matrix=pymupdf.Matrix(scale, scale),
        colorspace=pymupdf.csGRAY,
        alpha=False,
    )
    return Image.frombytes("L", (pixmap.width, pixmap.height), pixmap.samples)


def decode_page(
    page: pymupdf.Page,
    page_number: int,
    dpis: tuple[int, ...],
    cancel_event: threading.Event | None = None,
) -> MarkingCode:
    attempts: list[str] = []

    for dpi in dpis:
        if cancel_event is not None and cancel_event.is_set():
            raise ConversionCancelled("conversion cancelled")

        image = render_page(page, dpi)
        valid_results = [
            result for result in zxingcpp.read_barcodes(image) if result.valid
        ]
        gs1_results = [
            result
            for result in valid_results
            if result.symbology_identifier == "]d2"
        ]

        if len(gs1_results) == 1:
            return parse_honest_sign_code(gs1_results[0].bytes, page_number)

        attempts.append(
            f"{dpi} DPI: valid barcodes={len(valid_results)}, "
            f"GS1 DataMatrix={len(gs1_results)}"
        )

    raise ValueError("; ".join(attempts))


GS_CHAR = "\x1d"


def full_code_str(mc: MarkingCode) -> str:
    return mc.raw.decode("ascii")


def apply_escape_mode(value: str, escape_mode: str) -> str:
    # "ozon" keeps the real GS symbol (ASCII 29) — Ozon requires it.
    # "1c" replaces it with visible text \x1d.
    if escape_mode == "1c":
        return value.replace(GS_CHAR, r"\x1d")
    if escape_mode == "ozon":
        return value.replace(GS_CHAR, r"\u001d")
    return value


def write_codes_csv(
    path: Path,
    codes: list[MarkingCode],
    *,
    escape_csv: bool,
    escape_mode: str = "gs",
) -> None:
    path.parent.mkdir(parents=True, exist_ok=True)
    temporary_path: Path | None = None

    def transform(mc: MarkingCode) -> str:
        return apply_escape_mode(full_code_str(mc), escape_mode)

    try:
        with tempfile.NamedTemporaryFile(
            mode="w",
            encoding="utf-8-sig",
            newline="",
            prefix=f".{path.stem}_",
            suffix=".tmp",
            dir=path.parent,
            delete=False,
        ) as csv_file:
            temporary_path = Path(csv_file.name)
            if escape_csv:
                writer = csv.writer(csv_file, delimiter=",", quoting=csv.QUOTE_MINIMAL)
                writer.writerows([transform(mc)] for mc in codes)
            else:
                for mc in codes:
                    csv_file.write(transform(mc))
                    csv_file.write("\r\n")

        os.replace(temporary_path, path)
    finally:
        if temporary_path is not None and temporary_path.exists():
            temporary_path.unlink()


def write_codes_xlsx(
    path: Path,
    codes: list[MarkingCode],
    *,
    escape_mode: str = "gs",
) -> None:
    """Write all codes to a single .xlsx workbook with 10 sheets.

    Real GS (ASCII 29) cannot be stored inside an Excel cell (OpenXML forbids
    it), so in Excel the separator is written as visible text on every sheet:
    \x1d for 1C mode, \u001d for OZON mode.
    """
    import openpyxl
    from openpyxl.styles import Font, PatternFill

    sep = r"\u001d" if escape_mode == "ozon" else r"\x1d"

    def ex(mc: MarkingCode) -> str:
        return full_code_str(mc).replace(GS_CHAR, sep)

    full = [ex(mc) for mc in codes]
    gtin_list = [mc.gtin for mc in codes]
    serial_list = [mc.serial for mc in codes]
    key91_list = [mc.crypto_key_91 for mc in codes]
    crypto92_list = [mc.crypto_code_92 for mc in codes]
    ki = [full_code_str(mc).split(GS_CHAR, 1)[0] for mc in codes]
    header_fill = PatternFill("solid", fgColor="DDEBF7")
    header_font = Font(bold=True)

    gtin_counts: dict[str, int] = {}
    for g in gtin_list:
        gtin_counts[g] = gtin_counts.get(g, 0) + 1
    raw_seen: set[str] = set()
    dup_raw: set[str] = set()
    for mc in codes:
        if full_code_str(mc) in raw_seen:
            dup_raw.add(full_code_str(mc))
        raw_seen.add(full_code_str(mc))

    def make_sheet(ws, headers, rows) -> None:
        ws.append(headers)
        for c in ws[1]:
            c.font = header_font
            c.fill = header_fill
        for r in rows:
            ws.append(r)

    def crypto_ex(mc: MarkingCode) -> str:
        return f"{sep}91{mc.crypto_key_91}{sep}92{mc.crypto_code_92}"

    wb = openpyxl.Workbook()
    wb.remove(wb.active)

    # 1. Только полный код
    ws = wb.create_sheet("1. Полный код")
    make_sheet(ws, ["Код маркировки"], [[c] for c in full])

    # 2. GTIN + Серийный + Криптохвост
    ws = wb.create_sheet("2. GTIN+Серия+Хвост")
    make_sheet(
        ws,
        ["GTIN (01)", "Серийный (21)", "Криптохвост (91/92)"],
        [[g, s, crypto_ex(mc)] for g, s, mc in zip(gtin_list, serial_list, codes)],
    )

    # 3. Код идентификации
    ws = wb.create_sheet("3. Код идентификации")
    make_sheet(ws, ["Код идентификации (КИ)"], [[c] for c in ki])

    # 4. GTIN + Количество
    ws = wb.create_sheet("4. GTIN + Кол-во")
    make_sheet(ws, ["GTIN", "Кол-во"], [[g, n] for g, n in gtin_counts.items()])

    # 5. Все части по колонкам
    ws = wb.create_sheet("5. Части по колонкам")
    make_sheet(
        ws,
        ["GTIN", "Серийный", "Ключ 91", "Подпись 92"],
        [[g, s, k, c] for g, s, k, c in zip(gtin_list, serial_list, key91_list, crypto92_list)],
    )

    # 6. Полный код + Статус GTIN
    ws = wb.create_sheet("6. Код + Статус GTIN")
    make_sheet(
        ws,
        ["Код маркировки", "Статус GTIN"],
        [[c, ("OK" if gtin_check_digit_is_valid(g) else "Ошибка")]
         for c, g in zip(full, gtin_list)],
    )

    # 7. Полный код + Дубль
    ws = wb.create_sheet("7. Код + Дубль")
    make_sheet(
        ws,
        ["Код маркировки", "Дубль"],
        [[c, ("да" if full_code_str(mc) in dup_raw else "нет")] for c, mc in zip(full, codes)],
    )

    # 8. Для 1С (текст \x1d)
    ws = wb.create_sheet("8. Для 1С")
    make_sheet(ws, ["Код (для 1С)"], [[full_code_str(mc).replace(GS_CHAR, r"\x1d")] for mc in codes])

    # 9. Для OZON (текст \u001d)
    ws = wb.create_sheet("9. Для OZON")
    make_sheet(ws, ["Код (для OZON)"], [[full_code_str(mc).replace(GS_CHAR, r"\u001d")] for mc in codes])

    # 10. Номер + Полный код
    ws = wb.create_sheet("10. Номер + Код")
    make_sheet(ws, ["№", "Код маркировки"], [[i, c] for i, c in enumerate(full, start=1)])

    path.parent.mkdir(parents=True, exist_ok=True)
    wb.save(path)


def decode_pdf(
    pdf_path: Path,
    *,
    dpis: tuple[int, ...] = DEFAULT_DPIS,
    log: LogCallback | None = None,
    page_progress: PageProgressCallback | None = None,
    cancel_event: threading.Event | None = None,
) -> PdfDecodeResult:
    pdf_path = pdf_path.resolve()
    logger = log or (lambda _message: None)

    if not pdf_path.is_file():
        raise ConversionError(f"PDF file not found: {pdf_path}")
    if not dpis or any(dpi < 72 for dpi in dpis):
        raise ConversionError("DPI values must be >= 72")

    try:
        document = pymupdf.open(pdf_path)
    except Exception as exc:
        raise ConversionError(f"cannot open PDF: {exc}") from exc

    codes: list[MarkingCode] = []
    errors: list[str] = []

    with document:
        page_count = len(document)
        logger(f"PDF: {pdf_path}\n")
        logger(f"Pages: {page_count}\n")

        for page_index, page in enumerate(document):
            if cancel_event is not None and cancel_event.is_set():
                raise ConversionCancelled("conversion cancelled")

            page_number = page_index + 1
            try:
                codes.append(
                    decode_page(page, page_number, dpis, cancel_event=cancel_event)
                )
            except ConversionCancelled:
                raise
            except Exception as exc:
                errors.append(f"page {page_number}: {exc}")

            if page_progress is not None:
                page_progress(page_number, page_count)
            if page_number % 25 == 0 or page_number == page_count:
                logger(f"Processed: {page_number}/{page_count}\n")

    if errors:
        details = "\n".join(f"  {error}" for error in errors)
        raise ConversionError(
            f"conversion aborted; no CSV was written.\n{details}"
        )

    if cancel_event is not None and cancel_event.is_set():
        raise ConversionCancelled("conversion cancelled")

    return PdfDecodeResult(
        pdf_path=pdf_path,
        page_count=page_count,
        codes=tuple(codes),
    )


def convert_pdf(
    pdf_path: Path,
    csv_path: Path,
    *,
    dpis: tuple[int, ...] = DEFAULT_DPIS,
    log: LogCallback | None = None,
    page_progress: PageProgressCallback | None = None,
    cancel_event: threading.Event | None = None,
    escape_csv: bool = False,
) -> ConversionResult:
    csv_path = csv_path.resolve()
    logger = log or (lambda _message: None)
    decoded = decode_pdf(
        pdf_path,
        dpis=dpis,
        log=logger,
        page_progress=page_progress,
        cancel_event=cancel_event,
    )

    write_codes_csv(csv_path, list(decoded.codes), escape_csv=escape_csv)
    logger(f"Decoded: {decoded.code_count}\n")
    logger(f"Unique codes: {decoded.unique_code_count}\n")
    logger(f"GTIN counts: {decoded.gtins}\n")
    logger(f"CSV: {csv_path}\n")

    return ConversionResult(
        pdf_path=decoded.pdf_path,
        csv_path=csv_path,
        page_count=decoded.page_count,
        code_count=decoded.code_count,
        unique_code_count=decoded.unique_code_count,
        gtins=decoded.gtins,
    )


def find_zint_exe() -> Path | None:
    """Locate zint.exe (in _MEIPASS for PyInstaller, else alongside sources)."""
    candidates: list[Path] = []
    if getattr(sys, "frozen", False):
        candidates.append(Path(sys._MEIPASS) / "zint.exe")  # type: ignore[attr-defined]
        candidates.append(Path(sys.executable).resolve().parent / "zint.exe")
    candidates.append(Path(__file__).resolve().parent / "zint.exe")
    candidates += [
        Path("C:/Program Files/Zint/zint.exe"),
        Path("C:/Program Files (x86)/Zint/zint.exe"),
    ]
    for c in candidates:
        if c.is_file():
            return c
    return None


def parse_csv_codes(csv_text: str) -> list[str]:
    """Extract full marking codes from CSV text: one code per non-empty line.
    Comma is NOT treated as a separator (serial numbers may contain commas),
    so each physical line is read as a whole. Strips surrounding quotes and
    normalises GS separators to \\x1d."""
    import io

    lines = []
    for raw in io.StringIO(csv_text):
        cell = raw.rstrip("\r\n").strip()
        if not cell:
            continue
        # убираем обрамляющие кавычки, если есть
        if len(cell) >= 2 and cell[0] == '"' and cell[-1] == '"':
            cell = cell[1:-1]
        cell = cell.replace("\x1d", GS_CHAR).replace("\\x1d", GS_CHAR)
        if cell:
            lines.append(cell)
    return lines


def gs1data_to_zint_data(full_code: str) -> str:
    """Convert a raw Honest Sign code (with GS separators) into Zint GS1
    input using SQUARE brackets as AI delimiters: [01]GTIN[21]SER[91]KEY[92]SIG.

    Square brackets are Zint's default AI delimiters and do NOT conflict with
    round parentheses that may legally appear inside serial numbers (AI 21) —
    unlike `--gs1parens`, which misinterprets data parentheses as AI delimiters.
    """
    parts = full_code.split(GS_CHAR)
    gtin = ""
    serial = ""
    key91 = ""
    sig92 = ""
    rest = full_code
    if rest.startswith("01"):
        gtin = rest[2:16]
        rest = rest[16:]
    if rest.startswith("21"):
        serial = rest[2:]
        # serial ends at first GS
        if GS_CHAR in serial:
            serial = serial.split(GS_CHAR, 1)[0]
    if len(parts) >= 2:
        p91 = parts[1]
        if p91.startswith("91"):
            key91 = p91[2:]
    if len(parts) >= 3:
        p92 = parts[2]
        if p92.startswith("92"):
            sig92 = p92[2:]
    if not gtin:
        # fallback: assume the code already uses GS separators with AI prefixes
        out = []
        for part in full_code.split(GS_CHAR):
            out.append(part)
        return "".join("[" + p[:2] + "]" + p[2:] if len(p) > 2 else p for p in out)
    return f"[01]{gtin}[21]{serial}[91]{key91}[92]{sig92}"


def code_identification(full_code: str) -> str:
    """Возвращает укороченный код (01+GTIN+21+серия [GS] 91+ключ)
    из полного кода маркировки. Блок 91 берётся как есть, криптоблок 92
    отбрасывается."""
    parts = full_code.split(GS_CHAR)
    serial = ""
    rest = full_code
    if rest.startswith("01"):
        gtin = rest[2:16]
        rest = rest[16:]
    else:
        gtin = ""
    if rest.startswith("21"):
        serial = rest[2:]
        if GS_CHAR in serial:
            serial = serial.split(GS_CHAR, 1)[0]
    key91 = ""
    if len(parts) >= 2 and parts[1].startswith("91"):
        key91 = parts[1][2:]
    if key91:
        return f"01{gtin}21{serial}{GS_CHAR}91{key91}"
    return f"01{gtin}21{serial}"


def is_code_truncated(code: str) -> bool:
    """Проверяет, является ли код маркировки усечённым (содержит только
    код идентификации 01+GTIN+21+серия без криптохвоста 91/92).
    Такие коды не могут быть сгенерированы в полноценный DataMatrix."""
    text = code
    if text.startswith("01"):
        text = text[16:]
    if text.startswith("21"):
        text = text[2:]
    # после серии должен быть GS + 91 + ... — если нет, код обрезан
    return not (GS_CHAR in text and (GS_CHAR + "91") in text)


def gs1data_to_zint_from_short(short_code: str) -> str:
    """Строит Zint-данные для укороченного кода (01+GTIN+21+серия [GS] 91+ключ)
    в квадратных скобках — блок 91 берётся как есть, крипто 92 отбрасывается."""
    parts = short_code.split(GS_CHAR)
    gtin = ""
    serial = ""
    rest = short_code
    if rest.startswith("01"):
        gtin = rest[2:16]
        rest = rest[16:]
    if rest.startswith("21"):
        serial = rest[2:]
        if GS_CHAR in serial:
            serial = serial.split(GS_CHAR, 1)[0]
    key91 = ""
    if len(parts) >= 2 and parts[1].startswith("91"):
        key91 = parts[1][2:]
    if key91:
        return f"[01]{gtin}[21]{serial}[91]{key91}"
    return f"[01]{gtin}[21]{serial}"


def validate_code_structure(code: str) -> list[str]:
    """Проверяет структуру одного кода маркировки.
    Возвращает список найденных проблем (пустой — код корректен).
    Проверки: GTIN (14 цифр + контрольная), серия (AI 21), криптохвост
    (блоки 91 и 92)."""
    problems = []
    text = code
    if not text:
        return ["пустой код"]

    parts = text.split(GS_CHAR)
    # блок 01+GTIN
    gtin = ""
    rest = text
    if rest.startswith("01"):
        gtin = rest[2:16]
        rest = rest[16:]
    if len(gtin) != 14 or not gtin.isdigit():
        problems.append(f"GTIN должен быть 14 цифр, получено {len(gtin)}: {gtin!r}")
    elif not gtin_check_digit_is_valid(gtin):
        problems.append(f"неверная контрольная цифра GTIN: {gtin}")

    # серия (AI 21)
    if rest.startswith("21"):
        serial = rest[2:]
        if GS_CHAR in serial:
            serial = serial.split(GS_CHAR, 1)[0]
        if not serial:
            problems.append("пустой серийный номер (AI 21)")
        elif len(serial) > 20:
            problems.append(f"серийный номер слишком длинный: {len(serial)} симв. (макс. 20)")
    else:
        problems.append("нет серийного номера (AI 21)")

    # криптохвост: блоки 91 и 92
    block_texts = [p for p in parts if p.startswith(("91", "92"))]
    has91 = any(p.startswith("91") for p in block_texts)
    has92 = any(p.startswith("92") for p in block_texts)
    if not has91:
        problems.append("нет криптоключа (AI 91)")
    if not has92:
        problems.append("нет цифровой подписи (AI 92)")
    return problems


def validate_codes(codes: list[str]) -> list[dict]:
    """Проверяет список кодов маркировки и возвращает подробности по каждому.
    Каждый элемент: {"index", "code", "problems", "is_duplicate"}."""
    result = []
    seen: set[str] = set()
    for pos, code in enumerate(codes, 1):
        problems = validate_code_structure(code)
        dup = code in seen
        seen.add(code)
        if dup:
            problems.insert(0, "дубликат кода")
        result.append({
            "index": pos,
            "code": code,
            "problems": problems,
            "is_duplicate": dup,
        })
    return result


MM_PT = 72.0 / 25.4  # миллиметров в пунктах PDF


def build_label_page(
    doc: pymupdf.Document,
    zint_exe: Path,
    full_code: str,
    idx: int,
    *,
    label_w_mm: float,
    label_h_mm: float,
    heading: str = "",
    production_date: str = "",
    expiry: str = "",
    batch: str = "",
    show_date: bool = True,
    show_expiry: bool = True,
    show_batch: bool = True,
    show_code: bool = True,
    use_short: bool = False,
    scale: int = 4,
    bold_name: bool = False,
    bold_date: bool = False,
    bold_expiry: bool = False,
    bold_batch: bool = False,
    bold_code: bool = False,
    size_name: float | None = None,
    size_date: float | None = None,
    size_expiry: float | None = None,
    size_batch: float | None = None,
    size_code: float | None = None,
) -> None:
    """Рендерит одну этикетку нужного размера с DataMatrix по центру.

    use_short=True — DataMatrix строится из укороченного кода идентификации
    (01+GTIN+21+серия, без криптохвоста 91/92), а не из полного.
    Под штрихкодом печатается текст кода, дата производства, срок годности
    и номер партии — каждое поле независимо включается/выключается флагами
    show_date/show_expiry/show_batch/show_code.
    """
    import subprocess
    import tempfile

    page = doc.new_page(width=label_w_mm * MM_PT, height=label_h_mm * MM_PT)
    w_pt = label_w_mm * MM_PT
    h_pt = label_h_mm * MM_PT

    # пунктирная обводка границы этикетки (видно края наклейки при печати)
    try:
        page.draw_rect(
            pymupdf.Rect(0.2, 0.2, w_pt - 0.2, h_pt - 0.2),
            color=(0.6, 0.6, 0.6),
            width=0.4,
            dashes="[2 2] 0",
            fill=None,
        )
    except Exception:
        pass

    # встроенный шрифт с поддержкой кириллицы
    fontfile = ""
    arialbd = ""
    for candidate in (
        r"C:/Windows/Fonts/arial.ttf",
        r"C:/Windows/Fonts/times.ttf",
    ):
        if os.path.exists(candidate):
            fontfile = candidate
            break
    if os.path.exists(r"C:/Windows/Fonts/arialbd.ttf"):
        arialbd = r"C:/Windows/Fonts/arialbd.ttf"
    if fontfile:
        page.insert_font(fontname="F0", fontfile=fontfile)
        use_font = "F0"
    else:
        use_font = "helv"
    bold_font = "F1"
    try:
        if arialbd:
            page.insert_font(fontname=bold_font, fontfile=arialbd)
        else:
            bold_font = use_font
    except Exception:
        bold_font = use_font

    if use_short:
        zint_data = gs1data_to_zint_from_short(code_identification(full_code))
    else:
        zint_data = gs1data_to_zint_data(full_code)
    with tempfile.TemporaryDirectory() as td:
        png_path = Path(td) / f"{idx}.png"
        proc = subprocess.run(
            [
                str(zint_exe), "-b", "DATAMATRIX", "--gs1", "--gssep",
                "--square", f"--scale={scale}", "-d", zint_data,
                "-o", str(png_path),
            ],
            capture_output=True,
            text=True,
            creationflags=getattr(subprocess, "CREATE_NO_WINDOW", 0),
        )
        if proc.returncode != 0 or not png_path.exists():
            raise RuntimeError(
                f"Ошибка генерации DataMatrix для кода {idx + 1}: {proc.stderr.strip()}"
            )

        img = pymupdf.open(str(png_path))
        pix = img[0].get_pixmap()
        aspect = pix.width / pix.height

        top_margin = 14
        # нижняя часть — вспомогательный текст (наименование, дата, партия, код) —
        # отображается под кодом так же, как в макете и предпросмотре.
        bottom_lines: list[str] = []
        max_chars = 38 if w_pt > 200 else 30
        if heading:
            bottom_lines.extend(
                heading[i:i + max_chars]
                for i in range(0, len(heading), max_chars)
            )
        if show_date and production_date:
            bottom_lines.append(f"Дата производства: {production_date}")
        if show_expiry and expiry:
            bottom_lines.append(f"Срок годности: {expiry}")
        if show_batch and batch:
            bottom_lines.append(f"Партия: {batch}")
        if show_code:
            code_text = (code_identification(full_code) if use_short else full_code).replace(GS_CHAR, "\u29f8")
            wrapped = [
                code_text[i:i + max_chars]
                for i in range(0, len(code_text), max_chars)
            ] or [code_text]
            bottom_lines.extend(wrapped)

        # DataMatrix — сверху, крупный; текст (наименование, дата, партия) — ПОД кодом.
        import pymupdf as _mu
        # для точного измерения используем тот же шрифт, что и для печати (F0-Arial)
        _metric = _mu.Font(fontfile=fontfile) if fontfile else None

        def _tw(text, fs):
            if _metric is not None:
                return _metric.text_length(text, fontsize=fs)
            return _mu.get_text_length(text, fontname='helv', fontsize=fs)

        margin = 4  # отступ по краям этикетки, pt
        box_w = max(w_pt - 2 * margin, 8)

        # 1) определяем текстовую подпись кода (дублирует сам штрихкод — её можно
        #    отбросить первой при нехватке места; сам DataMatrix рисуем ВСЕГДА)
        code_text_lines: list[str] = []
        if show_code:
            code_text = (code_identification(full_code) if use_short else full_code).replace(GS_CHAR, "\u29f8")
            code_text_lines = [code_text[i:i + max_chars] for i in range(0, len(code_text), max_chars)] or [code_text]

        # остальной текст: дата, партия, наименование
        other_lines: list[str] = [ln for ln in bottom_lines if ln not in code_text_lines]

        # группа строк по типу (для порядка снизу вверх и приоритета отбрасывания)
        name_lines: list[str] = []
        date_lines: list[str] = []
        expiry_lines: list[str] = []
        batch_lines: list[str] = []
        for ln in other_lines:
            if ln.startswith("Дата производства:"):
                date_lines.append(ln)
            elif ln.startswith("Срок годности:"):
                expiry_lines.append(ln)
            elif ln.startswith("Партия:"):
                batch_lines.append(ln)
            elif ln.strip():
                name_lines.append(ln)
        # подпись кода всегда в самом низу (выкидывается первой при нехватке места)
        if show_code:
            code_with_sub = code_text_lines
        else:
            code_with_sub = []

        # ЕДИНЫЙ шрифт для даты/партии/наименования. Если задан ручной размер
        # (size_date и т.п.) — используем его; иначе автоподбор, чтобы дата и
        # партия влезали в ширину без переноса (минимум 3pt).
        def _pick(manual, strings):
            if manual and manual >= 2:
                return float(manual)
            fs = 6.5
            while fs > 3 and strings and any(_tw(s, fs) > box_w for s in strings):
                fs -= 0.5
            return fs

        info_fs_date = _pick(size_date, date_lines)
        info_fs_expiry = _pick(size_expiry, expiry_lines)
        info_fs_batch = _pick(size_batch, batch_lines)
        # если размеры заданы порознь и они различаются — это ок;
        # если нет — стараемся сделать общий размер для даты, срока и партии
        if size_date is None and size_expiry is None and size_batch is None:
            info_fs = min(info_fs_date, info_fs_expiry, info_fs_batch)
            info_fs_date = info_fs_expiry = info_fs_batch = info_fs
        name_fs = _pick(size_name, name_lines)
        code_fs = _pick(size_code, code_with_sub)

        def _wrap(text: str, ffs: float) -> list[tuple[str, float]]:
            if _tw(text, ffs) > box_w:
                chunks = []
                cur = ""
                for ch in text:
                    trial = cur + ch
                    if _tw(trial, ffs) > box_w and cur:
                        chunks.append(cur)
                        cur = ch
                    else:
                        cur = trial
                if cur:
                    chunks.append(cur)
                return [(c, ffs) for c in chunks]
            return [(text, ffs)]

        # дата, срок и партия — одна строка (без переноса) с общим шрифтом
        date_rows: list[tuple[str, float, str]] = [(t, info_fs_date, bold_font if bold_date else use_font) for t in date_lines]
        expiry_rows: list[tuple[str, float, str]] = [(t, info_fs_expiry, bold_font if bold_expiry else use_font) for t in expiry_lines]
        batch_rows: list[tuple[str, float, str]] = [(t, info_fs_batch, bold_font if bold_batch else use_font) for t in batch_lines]
        # наименование — тот же шрифт, но может переноситься на несколько строк
        name_rows: list[tuple[str, float, str]] = [(c, fs, bold_font if bold_name else use_font) for c, fs in [r for t in name_lines for r in _wrap(t, name_fs)]]
        # подпись кода — может переноситься
        code_rows: list[tuple[str, float, str]] = [(c, fs, bold_font if bold_code else use_font) for c, fs in [r for t in code_with_sub for r in _wrap(t, code_fs)]]

        line_h = 10  # межстрочный интервал между базовыми линиями (компактнее 12)
        margin = 4
        bottom_limit = h_pt - margin
        # Зазор между низом DataMatrix и первой строкой текста. Должен быть
        # заметно больше высоты надстрочной части шрифта (примерно 0.8*fs),
        # иначе верхние выносные буквы («Д» в «Дата», «П» в «Партия») визуально
        # наезжают на нижние ряды кода. Для fs вплоть до 8pt достаточно 8pt.
        gap_code_text = 8
        avail_w = w_pt - 2 * margin
        code_top = margin

        mandatory_date = date_rows  # всегда
        mandatory_expiry = expiry_rows  # всегда
        mandatory_batch = batch_rows  # всегда
        optional_name = name_rows
        optional_code = code_rows

        # сколько ещё строк поместится поверх даты, срока и партии
        block_mand = len(mandatory_date) + len(mandatory_expiry) + len(mandatory_batch)
        min_code_est = 18
        max_est_lines = max(int((bottom_limit - margin - min_code_est - gap_code_text) // line_h), 1)
        extra_slots = max_est_lines - block_mand

        # строим итоговый список сверху вниз: наименование, дата, срок, партия, подпись кода
        top_extra: list[tuple[str, float, str]] = []
        bottom_extra: list[tuple[str, float, str]] = []
        for item in optional_name:
            if len(top_extra) + len(bottom_extra) < extra_slots:
                top_extra.append(item)
        for item in optional_code:
            if len(top_extra) + len(bottom_extra) < extra_slots:
                bottom_extra.append(item)

        # итоговый порядок сверху вниз: наименование, дата, срок, партия, подпись кода
        selected: list[tuple[str, float, str]] = top_extra + mandatory_date + mandatory_expiry + mandatory_batch + bottom_extra

        # позиции базовых линий снизу вверх
        baseline_list: list[float] = []
        cur_y = bottom_limit
        for _ in range(len(selected)):
            baseline_list.append(cur_y - 3)
            cur_y -= line_h
        baseline_list.reverse()

        # код занимает всё свободное место от верха до начала текста
        first_text_top = baseline_list[0] if baseline_list else bottom_limit
        code_bottom = first_text_top - gap_code_text
        code_h = int(code_bottom - code_top)
        code_by_w = int(avail_w / aspect) if avail_w > 0 else code_h
        code_h = min(code_h, code_by_w)
        # Гарантия: низ изображения кода НИКОГДА не выходит за слот кода
        # (не пересекает текст). Даже если ограничение по ширине (code_by_w)
        # уменьшило высоту, низ кода остаётся выше первой строки текста.
        max_code_h = int(code_bottom - code_top)
        code_h = min(code_h, max_code_h)
        code_h = max(code_h, 0)
        code_w = max(int(code_h * aspect), 1)
        if code_h >= 6 and code_w >= 6:
            code_left = (w_pt - code_w) / 2
            page.insert_image(
                pymupdf.Rect(code_left, code_top, code_left + code_w, code_top + code_h),
                filename=str(png_path),
            )
        else:
            code_h = 0

        for ((chnk, ffs, fname), bl) in zip(selected, baseline_list):
            cw = _tw(chnk, ffs)
            page.insert_text(
                ((w_pt - cw) / 2, bl), chnk, fontsize=ffs, fontname=fname,
            )


def generate_label_pdf(
    csv_text: str,
    output_pdf: Path,
    *,
    label_w_mm: float = 58.0,
    label_h_mm: float = 40.0,
    heading: str = "",
    production_date: str = "",
    expiry: str = "",
    batch: str = "",
    show_date: bool = True,
    show_expiry: bool = True,
    show_batch: bool = True,
    show_code: bool = True,
    use_short: bool = False,
    scale: int = 4,
    bold_name: bool = False,
    bold_date: bool = False,
    bold_expiry: bool = False,
    bold_batch: bool = False,
    bold_code: bool = False,
    size_name: float | None = None,
    size_date: float | None = None,
    size_expiry: float | None = None,
    size_batch: float | None = None,
    size_code: float | None = None,
    progress_cb=None,
) -> tuple[int, str, Path]:
    """Генерирует PDF-этикетки с DataMatrix настраиваемого размера.

    Каждая этикетка — отдельная страница. Возвращает (page_count, error, save_path),
    где save_path — реальный путь, по которому был сохранён файл (учитывает
    механизм копий с номером вместо перезаписи).
    use_short=True — из полного кода строится укороченный (код идентификации,
    без криптохвоста 91/92), для перемаркировки.
    progress_cb(done, total) — вызывается после каждой этикетки для прогресс-бара.
    """
    import tempfile

    codes = parse_csv_codes(csv_text)
    if not codes:
        return 0, "Не найдено ни одного кода в CSV-файле", output_pdf

    zint_exe = find_zint_exe()
    if zint_exe is None:
        return 0, "Не найден zint.exe. Поместите его рядом с программой.", output_pdf

    import pymupdf

    skipped: list[str] = []
    ok_count = 0
    doc = pymupdf.open()
    try:
        for idx, full_code in enumerate(codes):
            try:
                build_label_page(
                    doc,
                    zint_exe,
                    full_code,
                    idx,
                    label_w_mm=label_w_mm,
                    label_h_mm=label_h_mm,
                    heading=heading,
                    production_date=production_date,
                    expiry=expiry,
                    batch=batch,
                    show_date=show_date,
                    show_expiry=show_expiry,
                    show_batch=show_batch,
                    show_code=show_code,
                    use_short=use_short,
                    scale=scale,
                    bold_name=bold_name,
                    bold_date=bold_date,
                    bold_expiry=bold_expiry,
                    bold_batch=bold_batch,
                    bold_code=bold_code,
                    size_name=size_name,
                    size_date=size_date,
                    size_expiry=size_expiry,
                    size_batch=size_batch,
                    size_code=size_code,
                )
                ok_count += 1
                if progress_cb:
                    progress_cb(ok_count, len(codes))
            except Exception as exc:
                # build_label_page успела создать пустую страницу — удаляем её
                if len(doc) > ok_count:
                    doc.delete_page(ok_count)
                skipped.append(f"код {idx + 1}: {exc}")
                if progress_cb:
                    progress_cb(idx + 1, len(codes))
    except Exception as exc:
        return 0, f"Ошибка: {exc}", output_pdf

    page_count = ok_count
    if page_count == 0:
        if skipped:
            detail = "\n".join(skipped)
            return 0, (f"Не удалось сгенерировать ни одной этикетки "
                       f"({len(skipped)} из {len(codes)} кодов пропущено):\n{detail}"), output_pdf
        return 0, "Не удалось сгенерировать ни одной этикетки", output_pdf

    output_pdf.parent.mkdir(parents=True, exist_ok=True)
    # Если файл с таким именем уже существует — создаём копию с номером
    # по порядку (имя_1.pdf, имя_2.pdf, ...), не перезаписывая старый.
    save_path = output_pdf
    n = 1
    while save_path.exists():
        save_path = output_pdf.with_stem(f"{output_pdf.stem}_{n}")
        n += 1
    save_note = ""
    if save_path != output_pdf:
        save_note = f"\nФайл уже существовал. Сохранено как: {save_path.name}"

    try:
        doc.save(str(save_path))
    except Exception as exc:
        return 0, f"Не удалось сохранить PDF: {exc}", output_pdf

    if skipped:
        detail = "\n".join(skipped)
        warning = (f"{len(skipped)} из {len(codes)} кодов пропущено "
                   f"(ошибочный формат):\n{detail}{save_note}")
        return page_count, warning, save_path
    return page_count, save_note, save_path


def generate_codes_pdf(
    csv_text: str,
    output_pdf: Path,
    *,
    label_text: bool = True,
    scale: int = 6,
) -> tuple[int, str]:
    """Generate a PDF where each code is on its own page, as GS1 DataMatrix.

    Returns (page_count, error_or_empty).
    """
    import subprocess
    import tempfile

    codes = parse_csv_codes(csv_text)
    if not codes:
        return 0, "Не найдено ни одного кода в CSV-файле"

    zint_exe = find_zint_exe()
    if zint_exe is None:
        return 0, "Не найден zint.exe. Поместите его рядом с программой."

    import pymupdf

    doc = pymupdf.open()
    try:
        for idx, full_code in enumerate(codes):
            zint_data = gs1data_to_zint_data(full_code)
            with tempfile.TemporaryDirectory() as td:
                png_path = Path(td) / f"{idx}.png"
                proc = subprocess.run(
                    [
                        str(zint_exe), "-b", "DATAMATRIX", "--gs1", "--gssep",
                        "--square", f"--scale={scale}", "-d", zint_data,
                        "-o", str(png_path),
                    ],
                    capture_output=True,
                    text=True,
                    creationflags=getattr(subprocess, "CREATE_NO_WINDOW", 0),
                )
                if proc.returncode != 0 or not png_path.exists():
                    return 0, f"Ошибка генерации DataMatrix для кода {idx + 1}: {proc.stderr.strip()}"

                page = doc.new_page(width=595, height=842)
                page.insert_image(pymupdf.Rect(197, 220, 398, 421), filename=str(png_path))
                if label_text:
                    y = 500
                    display = full_code.replace(GS_CHAR, "\u29f8")
                    for j in range(0, len(display), 38):
                        page.insert_text((135, y), display[j:j + 38], fontsize=7, fontname="helv")
                        y += 12
    except Exception as exc:
        return 0, f"Ошибка: {exc}"
    finally:
        pass

    output_pdf.parent.mkdir(parents=True, exist_ok=True)
    doc.save(str(output_pdf))
    return doc.page_count, ""


def export_labels_png(
    csv_text: str,
    output_dir: Path,
    *,
    label_w_mm: float = 58.0,
    label_h_mm: float = 40.0,
    heading: str = "",
    production_date: str = "",
    expiry: str = "",
    batch: str = "",
    show_date: bool = True,
    show_expiry: bool = True,
    show_batch: bool = True,
    show_code: bool = True,
    use_short: bool = False,
    dpi: int = 300,
    bold_name: bool = False,
    bold_date: bool = False,
    bold_expiry: bool = False,
    bold_batch: bool = False,
    bold_code: bool = False,
    size_name: float | None = None,
    size_date: float | None = None,
    size_expiry: float | None = None,
    size_batch: float | None = None,
    size_code: float | None = None,
) -> tuple[int, str]:
    """Экспортирует каждую этикетку как отдельный PNG-файл.

    Возвращает (число созданных файлов, ошибка_или_пусто). Каждая этикетка
    отрисовывается из PDF-страницы с указанным DPI.
    """
    from PIL import Image

    tmp_pdf = output_dir / "__delfin_tmp.pdf"
    # Удаляем устаревший временный PDF, чтобы generate_label_pdf сохранил
    # файл ровно под этим именем (без суффикса _1, _2) и вернул его в save_path.
    try:
        if tmp_pdf.exists():
            tmp_pdf.unlink()
    except OSError:
        pass
    pages, err, save_path = generate_label_pdf(
        csv_text,
        tmp_pdf,
        label_w_mm=label_w_mm,
        label_h_mm=label_h_mm,
        heading=heading,
        production_date=production_date,
        expiry=expiry,
        batch=batch,
        show_date=show_date,
        show_expiry=show_expiry,
        show_batch=show_batch,
        show_code=show_code,
        use_short=use_short,
        scale=6,
        bold_name=bold_name,
        bold_date=bold_date,
        bold_expiry=bold_expiry,
        bold_batch=bold_batch,
        bold_code=bold_code,
        size_name=size_name,
        size_date=size_date,
        size_expiry=size_expiry,
        size_batch=size_batch,
        size_code=size_code,
    )
    if err and pages == 0:
        return 0, err

    import pymupdf

    output_dir.mkdir(parents=True, exist_ok=True)
    # открываем именно тот файл, который реально создан (мог быть суффикс _1, _2)
    rendered_pdf = save_path if save_path is not None and save_path.exists() else tmp_pdf
    doc = pymupdf.open(str(rendered_pdf))
    saved = 0
    try:
        for i in range(len(doc)):
            pix = doc[i].get_pixmap(dpi=dpi)
            img = Image.frombytes("RGB", (pix.width, pix.height), pix.samples)
            out_file = output_dir / f"etiketa_{i + 1:04d}.png"
            img.save(str(out_file))
            saved += 1
    finally:
        doc.close()
        # удаляем и временный PDF (если сохранялся под исходным именем),
        # и возможную копию с номером (save_path вроде __delfin_tmp_1.pdf)
        for candidate in {tmp_pdf, rendered_pdf}:
            try:
                candidate.unlink()
            except OSError:
                pass
    return saved, ""


def print_pdf_windows(pdf_path: Path) -> str:
    """Отправляет PDF на печать штатным диалогом Windows.

    Возвращает пустую строку при успехе или текст ошибки.
    """
    if not pdf_path.is_file():
        return f"PDF-файл не найден: {pdf_path}"
    try:
        os.startfile(str(pdf_path), "print")  # type: ignore[attr-defined]
        return ""
    except Exception as exc:
        return f"Не удалось запустить печать: {exc}"

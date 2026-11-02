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
    Strips surrounding quotes, and normalises GS separators to \\x1d."""
    import csv as _csv
    import io

    lines = []
    reader = _csv.reader(io.StringIO(csv_text))
    for row in reader:
        if row:
            cell = row[0].strip()
            if cell:
                cell = cell.replace("\x1d", GS_CHAR).replace("\\x1d", GS_CHAR).replace("è", GS_CHAR)
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
    """Возвращает укороченный код идентификации (01+GTIN+21+серия)
    из полного кода маркировки. Без криптохвоста (91/92)."""
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
    """Строит Zint-данные для укороченного кода (01+GTIN+21+серия)
    в квадратных скобках — без криптохвоста."""
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
    return f"[01]{gtin}[21]{serial}"


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
    batch: str = "",
    show_date: bool = True,
    show_batch: bool = True,
    show_code: bool = True,
    use_short: bool = False,
    scale: int = 4,
) -> None:
    """Рендерит одну этикетку нужного размера с DataMatrix по центру.

    use_short=True — DataMatrix строится из укороченного кода идентификации
    (01+GTIN+21+серия, без криптохвоста 91/92), а не из полного.
    Под штрихкодом печатается текст кода, дата производства и номер партии —
    каждое поле независимо включается/выключается флагами show_date/show_batch/show_code.
    """
    import subprocess
    import tempfile

    page = doc.new_page(width=label_w_mm * MM_PT, height=label_h_mm * MM_PT)
    w_pt = label_w_mm * MM_PT
    h_pt = label_h_mm * MM_PT

    # встроенный шрифт с поддержкой кириллицы
    fontfile = ""
    for candidate in (
        r"C:/Windows/Fonts/arial.ttf",
        r"C:/Windows/Fonts/times.ttf",
    ):
        if os.path.exists(candidate):
            fontfile = candidate
            break
    if fontfile:
        page.insert_font(fontname="F0", fontfile=fontfile)
        use_font = "F0"
    else:
        use_font = "helv"

    if use_short:
        zint_data = gs1data_to_zint_from_short(code_identification(full_code))
    else:
        zint_data = gs1data_to_zint_data(full_code)
    with tempfile.TemporaryDirectory() as td:
        png_path = Path(td) / f"{idx}.png"
        proc = subprocess.run(
            [
                str(zint_exe), "-b", "DATAMATRIX", "--gs1", "--gssep",
                f"--scale={scale}", "-d", zint_data,
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
        if show_batch and batch:
            bottom_lines.append(f"Партия: {batch}")
        if show_code:
            code_text = (code_identification(full_code) if use_short else full_code).replace(GS_CHAR, "\u29f8")
            wrapped = [
                code_text[i:i + max_chars]
                for i in range(0, len(code_text), max_chars)
            ] or [code_text]
            bottom_lines.extend(wrapped)

        # DataMatrix — в верхней части этикетки (крупный), текст под ним.
        # Код опущен ниже от верхнего края, чтобы при смещении листа в принтере
        # не смазаться по краю и остаться на этикетке.
        code_top = top_margin + 24
        text_h = len(bottom_lines) * 11 + 8
        avail_h = (h_pt - 10) - code_top - text_h
        if avail_h < 8:
            avail_h = 8
        avail_w = w_pt - 16
        code_h = min(int(avail_h), int(avail_w / aspect))
        code_w = int(code_h * aspect)

        code_left = (w_pt - code_w) / 2
        page.insert_image(
            pymupdf.Rect(code_left, code_top, code_left + code_w, code_top + code_h),
            filename=str(png_path),
        )

        # текст под штрихкодом — по центру этикетки
        y = code_top + code_h + 6
        for line in bottom_lines:
            fontsize = 7 if len(line) > 30 else 8
            line_h = fontsize + 4
            page.insert_textbox(
                pymupdf.Rect(8, y, w_pt - 8, y + line_h),
                line,
                fontsize=fontsize,
                fontname=use_font,
                align=1,
            )
            y += line_h

        # подпись номера этикетки — в самом низу этикетки
        try:
            page.insert_text(
                (6, h_pt - 3),
                f"Этикетка {idx + 1}",
                fontsize=6,
                fontname=use_font,
            )
        except Exception:
            pass


def generate_label_pdf(
    csv_text: str,
    output_pdf: Path,
    *,
    label_w_mm: float = 58.0,
    label_h_mm: float = 40.0,
    heading: str = "",
    production_date: str = "",
    batch: str = "",
    show_date: bool = True,
    show_batch: bool = True,
    show_code: bool = True,
    use_short: bool = False,
    scale: int = 4,
) -> tuple[int, str]:
    """Генерирует PDF-этикетки с DataMatrix настраиваемого размера.

    Каждая этикетка — отдельная страница. Возвращает (page_count, error).
    use_short=True — из полного кода строится укороченный (код идентификации,
    без криптохвоста 91/92), для перемаркировки.
    """
    import tempfile

    codes = parse_csv_codes(csv_text)
    if not codes:
        return 0, "Не найдено ни одного кода в CSV-файле"

    zint_exe = find_zint_exe()
    if zint_exe is None:
        return 0, "Не найден zint.exe. Поместите его рядом с программой."

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
                    batch=batch,
                    show_date=show_date,
                    show_batch=show_batch,
                    show_code=show_code,
                    use_short=use_short,
                    scale=scale,
                )
                ok_count += 1
            except Exception as exc:
                # build_label_page успела создать пустую страницу — удаляем её
                if len(doc) > ok_count:
                    doc.delete_page(ok_count)
                skipped.append(f"код {idx + 1}: {exc}")
    except Exception as exc:
        return 0, f"Ошибка: {exc}"

    page_count = ok_count
    if page_count == 0:
        if skipped:
            detail = "\n".join(skipped)
            return 0, (f"Не удалось сгенерировать ни одной этикетки "
                       f"({len(skipped)} из {len(codes)} кодов пропущено):\n{detail}")
        return 0, "Не удалось сгенерировать ни одной этикетки"

    output_pdf.parent.mkdir(parents=True, exist_ok=True)
    doc.save(str(output_pdf))

    if skipped:
        detail = "\n".join(skipped)
        warning = (f"{len(skipped)} из {len(codes)} кодов пропущено "
                   f"(ошибочный формат):\n{detail}")
        return page_count, warning
    return page_count, ""


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
                        f"--scale={scale}", "-d", zint_data,
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
    batch: str = "",
    show_date: bool = True,
    show_batch: bool = True,
    show_code: bool = True,
    use_short: bool = False,
    dpi: int = 300,
) -> tuple[int, str]:
    """Экспортирует каждую этикетку как отдельный PNG-файл.

    Возвращает (число созданных файлов, ошибка_или_пусто). Каждая этикетка
    отрисовывается из PDF-страницы с указанным DPI.
    """
    from PIL import Image

    tmp_pdf = output_dir / "__delfin_tmp.pdf"
    pages, err = generate_label_pdf(
        csv_text,
        tmp_pdf,
        label_w_mm=label_w_mm,
        label_h_mm=label_h_mm,
        heading=heading,
        production_date=production_date,
        batch=batch,
        show_date=show_date,
        show_batch=show_batch,
        show_code=show_code,
        use_short=use_short,
        scale=6,
    )
    if err and pages == 0:
        return 0, err

    import pymupdf

    output_dir.mkdir(parents=True, exist_ok=True)
    doc = pymupdf.open(str(tmp_pdf))
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
        try:
            tmp_pdf.unlink()
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

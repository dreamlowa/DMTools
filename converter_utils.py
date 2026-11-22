"""Чистая логика конвертации файлов и вспомогательные функции для печати
текста в PDF с кириллицей. Выделено из gui.py, чтобы не раздувать GUI-код.

Не импортирует tkinter и не зависит от GUI — только от стандартных библиотек
и установленных зависимостей (PyMuPDF, python-docx, openpyxl, markdown, PIL).
"""
from __future__ import annotations

import os
from pathlib import Path

# Путь к файлу шрифта для кириллицы (кэш).
_PDF_GUI_FONT_FILE: str | None = None
_PDF_GUI_FONT_NAME: str | None = None


def _pdf_text_page_font(page) -> str:
    """Возвращает имя кириллического шрифта, зарегистрированного в странице PDF.

    page.insert_text без шрифта использует Helvetica, в которой нет кириллицы,
    и русский текст выводится как «···». Здесь подключаем системный шрифт
    (Arial/Times) с кириллицей и регистрируем его в текущей странице через
    insert_font, возвращая имя для использования в insert_text/get_text_length.
    """
    import pymupdf as _mu
    global _PDF_GUI_FONT_NAME, _PDF_GUI_FONT_FILE
    if _PDF_GUI_FONT_NAME is None:
        fontfile = None
        for cand in (
            r"C:/Windows/Fonts/arial.ttf",
            r"C:/Windows/Fonts/times.ttf",
            r"C:/Windows/Fonts/DejaVuSans.ttf",
        ):
            if os.path.exists(cand):
                fontfile = cand
                break
        if fontfile:
            _PDF_GUI_FONT_NAME = fontfile
            _PDF_GUI_FONT_FILE = fontfile
        else:
            _PDF_GUI_FONT_NAME = "helv"
    if _PDF_GUI_FONT_NAME != "helv":
        try:
            page.insert_font(fontname="F_cp", fontfile=_PDF_GUI_FONT_NAME)
            return "F_cp"
        except Exception:
            pass
    return "helv"


def _pdf_wrap_text(page, text: str, x: float, y: float, size: float,
                   max_width: float, line_h: float) -> float:
    """Вставляет текст в PDF с кириллическим шрифтом и переносом по ширине.

    Возвращает новую координату y после отрисовки всех перенесённых строк.
    """
    import pymupdf as _mu
    fontname = _pdf_text_page_font(page)
    measurer = None
    global _PDF_GUI_FONT_FILE
    if fontname == "F_cp" and _PDF_GUI_FONT_FILE is not None:
        try:
            measurer = _mu.Font(fontfile=_PDF_GUI_FONT_FILE)
        except Exception:
            measurer = None
    lines: list[str] = []
    for raw in text.splitlines():
        if not raw:
            lines.append("")
            continue
        # вставка через Arial кодирует обычный пробел как \xa0 — нормализуем
        raw = raw.replace("\xa0", " ")
        current = ""
        for ch in raw:
            trial = current + ch
            if measurer is not None:
                width = measurer.text_length(trial, fontsize=size)
            else:
                width = _mu.get_text_length(trial, fontname=fontname, fontsize=size)
            if width > max_width and current:
                lines.append(current)
                current = ch
            else:
                current = trial
        if current:
            lines.append(current)
    for line in lines:
        page.insert_text((x, y), line, fontsize=size, fontname=fontname)
        y += line_h
    return y


# Метки форматов, доступные в выпадающих списках конвертера.
FORMAT_LABELS = ("PDF (.pdf)", "Word (.docx)", "HTML (.html)", "TXT (.txt)",
                 "Markdown (.md)", "CSV (.csv)", "Excel (.xlsx)", "PNG (.png)",
                 "JPG (.jpg)")

# Реальные поддерживаемые конверсии: "Из формата" -> список "В формат".
CONV_SUPPORTED = {
    "PDF (.pdf)": ["Word (.docx)", "TXT (.txt)", "Excel (.xlsx)"],
    "Word (.docx)": ["PDF (.pdf)"],
    "HTML (.html)": ["PDF (.pdf)", "Word (.docx)", "TXT (.txt)"],
    "TXT (.txt)": ["PDF (.pdf)", "Word (.docx)", "HTML (.html)"],
    "Markdown (.md)": ["HTML (.html)", "PDF (.pdf)", "Word (.docx)", "TXT (.txt)"],
    "CSV (.csv)": ["Excel (.xlsx)", "Word (.docx)", "TXT (.txt)", "HTML (.html)", "PDF (.pdf)"],
    "Excel (.xlsx)": ["CSV (.csv)", "Word (.docx)", "TXT (.txt)", "HTML (.html)", "PDF (.pdf)"],
    "PNG (.png)": ["PNG (.png)", "JPG (.jpg)", "PDF (.pdf)"],
    "JPG (.jpg)": ["PNG (.png)", "JPG (.jpg)", "PDF (.pdf)"],
}

_SRC_EXT = {
    "PDF (.pdf)": ".pdf", "Word (.docx)": ".docx", "HTML (.html)": ".html",
    "TXT (.txt)": ".txt", "Markdown (.md)": ".md", "CSV (.csv)": ".csv",
    "Excel (.xlsx)": ".xlsx", "PNG (.png)": ".png", "JPG (.jpg)": ".jpg",
}
_DST_EXT = {
    "Word (.docx)": ".docx", "PDF (.pdf)": ".pdf", "TXT (.txt)": ".txt",
    "HTML (.html)": ".html", "Excel (.xlsx)": ".xlsx", "CSV (.csv)": ".csv",
    "PNG (.png)": ".png", "JPG (.jpg)": ".jpg",
}


def ext_from_src_label(label: str) -> str:
    return _SRC_EXT.get(label, "")


def ext_from_dst_label(label: str) -> str:
    return _DST_EXT.get(label, ".out")


def unique_target(src: Path, to: str) -> Path:
    """Возвращает уникальное имя файла результата (с номером, если уже есть)."""
    ext = ext_from_dst_label(to)
    target = src.with_suffix(ext)
    n = 1
    base = target
    while target.exists():
        target = base.with_stem(f"{base.stem}_{n}")
        n += 1
    return target


def convert(frm: str, to: str, src_path: Path, target: Path,
            table_cell_text=None) -> None:
    """Выполняет конкретную конверсию. Вызывает исключение при ошибке.

    table_cell_text — необязательный callable(page, bbox)->str, используемый
    для извлечения текста ячейки PDF->Excel (передаётся из GUI, чтобы не
    дублировать логику).
    """
    import fitz  # PyMuPDF

    norm_frm = frm.split(" ")[0] if frm else frm
    frm = norm_frm

    if frm == "PDF" and to == "Word (.docx)":
        from docx import Document
        doc = Document()
        with fitz.open(str(src_path)) as pdf:
            for page in pdf:
                text = page.get_text("text")
                if not text.strip():
                    continue
                for line in text.splitlines():
                    doc.add_paragraph(line.replace("\xa0", " "))
        doc.save(str(target))
    elif frm == "PDF" and to == "TXT (.txt)":
        with fitz.open(str(src_path)) as pdf:
            text = "\n".join(page.get_text("text") for page in pdf)
        target.write_text(text.replace("\xa0", " "), encoding="utf-8")
    elif frm == "Word" and to == "PDF (.pdf)":
        from docx import Document as D
        doc = D(str(src_path))
        out = fitz.open()
        page = out.new_page()
        y = 50
        for para in doc.paragraphs:
            t = para.text.strip()
            if not t:
                continue
            y = _pdf_wrap_text(page, t, 50, y, 11, 490, 18)
            if y > 790:
                page = out.new_page()
                y = 50
        out.save(str(target))
        out.close()
    elif frm == "HTML" and to == "PDF (.pdf)":
        html = src_path.read_text(encoding="utf-8", errors="replace")
        out = fitz.open()
        page = out.new_page()
        page.insert_htmlbox(fitz.Rect(50, 50, 545, 800), html)
        out.save(str(target))
        out.close()
    elif frm == "HTML" and to == "Word (.docx)":
        from docx import Document as D
        import html as _html
        import re
        html = src_path.read_text(encoding="utf-8", errors="replace")
        doc = D()
        for block in re.split(r"<\s*(?:p|div|li|h[1-6]|br)[^>]*>", html):
            txt = re.sub(r"<[^>]+>", "", block)
            txt = _html.unescape(txt).strip()
            if txt:
                doc.add_paragraph(txt)
        doc.save(str(target))
    elif frm == "TXT" and to == "PDF (.pdf)":
        text = src_path.read_text(encoding="utf-8", errors="replace")
        out = fitz.open()
        page = out.new_page()
        y = 50
        for line in text.splitlines():
            y = _pdf_wrap_text(page, line, 50, y, 11, 490, 18)
            if y > 790:
                page = out.new_page()
                y = 50
        out.save(str(target))
        out.close()
    elif frm == "TXT" and to == "Word (.docx)":
        from docx import Document as D
        doc = D()
        for line in src_path.read_text(encoding="utf-8", errors="replace").splitlines():
            doc.add_paragraph(line)
        doc.save(str(target))
    elif frm == "Markdown" and to == "HTML (.html)":
        import markdown as md
        html = md.markdown(src_path.read_text(encoding="utf-8", errors="replace"))
        target.write_text(html, encoding="utf-8")
    elif frm == "Markdown" and to == "PDF (.pdf)":
        import markdown as md
        html = "<html><body>" + md.markdown(
            src_path.read_text(encoding="utf-8", errors="replace")) + "</body></html>"
        out = fitz.open()
        page = out.new_page()
        page.insert_htmlbox(fitz.Rect(50, 50, 545, 800), html)
        out.save(str(target))
        out.close()
    elif frm == "CSV" and to == "Excel (.xlsx)":
        import csv
        import openpyxl
        wb = openpyxl.Workbook()
        ws = wb.active
        with src_path.open("r", encoding="utf-8", errors="replace", newline="") as f:
            for row in csv.reader(f):
                ws.append(row)
        wb.save(str(target))
    elif frm == "Excel" and to == "CSV (.csv)":
        import csv
        import openpyxl
        wb = openpyxl.load_workbook(str(src_path), read_only=True)
        ws = wb.active
        with target.open("w", encoding="utf-8", newline="") as f:
            writer = csv.writer(f)
            for row in ws.iter_rows(values_only=True):
                writer.writerow(row)
        wb.close()
    elif frm in ("PNG", "JPG") and to in ("PNG (.png)", "JPG (.jpg)"):
        from PIL import Image
        img = Image.open(str(src_path))
        if to == "JPG (.jpg)":
            img = img.convert("RGB")
            img.save(str(target), "JPEG", quality=90)
        else:
            img.save(str(target), "PNG")
    elif frm in ("PNG", "JPG") and to == "PDF (.pdf)":
        from PIL import Image
        img = Image.open(str(src_path)).convert("RGB")
        img.save(str(target), "PDF", resolution=150.0)
    elif frm == "PDF" and to == "Excel (.xlsx)":
        import openpyxl
        from openpyxl.styles import Border, Side
        thin = Side(style="thin", color="000000")
        border = Border(left=thin, right=thin, top=thin, bottom=thin)

        wb = openpyxl.Workbook()
        wb.remove(wb.active)
        page_no = 0
        with fitz.open(str(src_path)) as pdf:
            for page in pdf:
                page_no += 1
                page_w = page.rect.width
                page_h = page.rect.height
                tables = page.find_tables()
                if not tables.tables:
                    ws = wb.create_sheet(title=f"Стр. {page_no}")
                    scale_x = 70
                    scale_y = 32
                    for block in page.get_text("dict")["blocks"]:
                        if block.get("type") != 0:
                            continue
                        for line in block.get("lines", []):
                            x0 = line.get("bbox", [0, 0, 0, 0])[0]
                            y0 = line.get("bbox", [0, 0, 0, 0])[1]
                            col = max(1, int(round(x0 / page_w * scale_x)) + 1)
                            row = max(1, int(round(y0 / page_h * scale_y)) + 1)
                            text = "".join(
                                span.get("text", "") for span in line.get("spans", []))
                            if text.strip():
                                ws.cell(row=row, column=int(col), value=text)
                    continue
                for t_idx, table in enumerate(tables.tables, start=1):
                    ws = wb.create_sheet(title=f"Стр. {page_no}.{t_idx}")
                    data = []
                    for row_obj in table.rows:
                        row_vals = []
                        for bbox in row_obj.cells:
                            if bbox is None:
                                row_vals.append("")
                                continue
                            if table_cell_text is not None:
                                row_vals.append(table_cell_text(page, bbox))
                            else:
                                row_vals.append("")
                        data.append(row_vals)
                    for ri, row_vals in enumerate(data):
                        for ci, val in enumerate(row_vals):
                            if val is None or not str(val).strip():
                                continue
                            cell = ws.cell(row=ri + 1, column=ci + 1, value=val)
                            cell.border = border
        if not wb.sheetnames:
            wb.create_sheet(title="Лист 1")
        wb.save(str(target))
    elif frm in ("CSV", "Excel") and to in ("Word (.docx)", "TXT (.txt)", "HTML (.html)", "PDF (.pdf)"):
        import csv as _csv
        rows = []
        if frm == "CSV":
            with src_path.open("r", encoding="utf-8", errors="replace", newline="") as f:
                rows = [row for row in _csv.reader(f)]
        else:
            import openpyxl
            wb = openpyxl.load_workbook(str(src_path), read_only=True)
            ws = wb.active
            rows = [list(r) for r in ws.iter_rows(values_only=True)]
            wb.close()
        text_lines = ["\t".join(("" if c is None else str(c)) for c in r) for r in rows]
        if to == "TXT (.txt)":
            target.write_text("\n".join(text_lines), encoding="utf-8")
        elif to == "HTML (.html)":
            import html as _h
            body = "\n".join(
                "<tr>" + "".join(f"<td>{_h.escape('' if c is None else str(c))}</td>" for c in r) + "</tr>"
                for r in rows)
            target.write_text(
                "<html><body><table>" + body + "</table></body></html>",
                encoding="utf-8")
        elif to == "Word (.docx)":
            from docx import Document as _D
            doc = _D()
            for r in rows:
                doc.add_paragraph("\t".join(("" if c is None else str(c)) for c in r))
            doc.save(str(target))
        else:  # PDF
            out = fitz.open()
            page = out.new_page()
            y = 50
            for line in text_lines:
                y = _pdf_wrap_text(page, line, 50, y, 11, 490, 18)
                if y > 790:
                    page = out.new_page()
                    y = 50
            out.save(str(target))
            out.close()
    elif frm == "TXT" and to == "HTML (.html)":
        import html as _h
        text = src_path.read_text(encoding="utf-8", errors="replace")
        paras = "".join(f"<p>{_h.escape(ln)}</p>" for ln in text.splitlines() if ln.strip())
        target.write_text("<html><body>" + paras + "</body></html>", encoding="utf-8")
    elif frm == "Markdown" and to in ("Word (.docx)", "TXT (.txt)"):
        import re
        raw = src_path.read_text(encoding="utf-8", errors="replace")
        lines = []
        for ln in raw.splitlines():
            t = re.sub(r"^#{1,6}\s*", "", ln)
            t = re.sub(r"\*\*([^*]+)\*\*", r"\1", t)
            t = re.sub(r"\*([^*]+)\*", r"\1", t)
            t = re.sub(r"\[([^\]]+)\]\([^)]*\)", r"\1", t)
            t = re.sub(r"^[-*]\s+", "", t)
            t = re.sub(r"^\d+\.\s+", "", t)
            t = re.sub(r"`([^`]+)`", r"\1", t)
            t = re.sub(r"!\[[^\]]*\]\([^)]*\)", "", t)
            if t.strip():
                lines.append(t.strip())
        if to == "TXT (.txt)":
            target.write_text("\n".join(lines), encoding="utf-8")
        else:
            from docx import Document as _D
            doc = _D()
            for ln in lines:
                doc.add_paragraph(ln)
            doc.save(str(target))
    elif frm == "HTML" and to == "TXT (.txt)":
        import re
        import html as _h
        html = src_path.read_text(encoding="utf-8", errors="replace")
        text = re.sub(r"<[^>]+>", "", html)
        text = _h.unescape(text)
        target.write_text(text.strip() + "\n", encoding="utf-8")
    else:
        raise RuntimeError(f"Конверсия {frm} → {to} пока не поддерживается.")

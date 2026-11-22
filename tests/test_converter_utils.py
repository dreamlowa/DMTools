"""Тесты для converter_utils: проверка всех поддерживаемых конверсий файлов
и вспомогательных функций (ext_from_*, unique_target, CONV_SUPPORTED).

Запуск: python -m unittest discover -s tests -v
"""
from __future__ import annotations

import tempfile
import unittest
from pathlib import Path

import openpyxl

import converter_utils as cu


def _make_pdf(path: Path, text: str = "Привет мир") -> None:
    import pymupdf
    font = r"C:/Windows/Fonts/arial.ttf"
    d = pymupdf.open()
    page = d.new_page()
    if Path(font).exists():
        page.insert_font(fontname="Fc", fontfile=font)
        page.insert_text((50, 50), text, fontname="Fc")
    else:
        page.insert_text((50, 50), "hello world")
    d.save(str(path))


class ExtMappingTests(unittest.TestCase):
    def test_src_ext(self) -> None:
        self.assertEqual(cu.ext_from_src_label("PDF (.pdf)"), ".pdf")
        self.assertEqual(cu.ext_from_src_label("Excel (.xlsx)"), ".xlsx")
        self.assertEqual(cu.ext_from_src_label("неизвестный"), "")

    def test_dst_ext(self) -> None:
        self.assertEqual(cu.ext_from_dst_label("Word (.docx)"), ".docx")
        self.assertEqual(cu.ext_from_dst_label("JPG (.jpg)"), ".jpg")
        self.assertEqual(cu.ext_from_dst_label("???"), ".out")

    def test_unique_target(self) -> None:
        with tempfile.TemporaryDirectory() as td:
            src = Path(td) / "doc.pdf"
            src.touch()
            t1 = cu.unique_target(src, "TXT (.txt)")
            self.assertEqual(t1.name, "doc.txt")
            t1.touch()
            t2 = cu.unique_target(src, "TXT (.txt)")
            self.assertEqual(t2.name, "doc_1.txt")

    def test_supported_is_symmetric_consistent(self) -> None:
        # каждый целевой формат — также валидная метка исходного формата
        # (кроме изображений, где разрешена ре-кодировка PNG->PNG / JPG->JPG)
        image_formats = {"PNG (.png)", "JPG (.jpg)"}
        for frm, targets in cu.CONV_SUPPORTED.items():
            for tgt in targets:
                self.assertIn(tgt, cu.CONV_SUPPORTED,
                              f"цель {tgt} не существует в карте")
            if frm not in image_formats:
                self.assertNotIn(frm, targets,
                                 f"исходный и целевой совпадают для {frm}")


class ConversionTest(unittest.TestCase):
    def _conv(self, frm: str, to: str, src: Path, td: Path) -> Path:
        target = cu.unique_target(src, to)
        cu.convert(frm, to, src, target)
        return target

    def test_all_supported_pairs_produce_output(self) -> None:
        with tempfile.TemporaryDirectory() as temporary_directory:
            td = Path(temporary_directory)
            # готовим исходники
            src = {}
            pdf = td / "s.pdf"
            _make_pdf(pdf, "Привет")
            src["PDF (.pdf)"] = pdf
            from docx import Document
            doc = Document()
            doc.add_paragraph("Привет")
            w2 = td / "s.docx"
            doc.save(str(w2))
            src["Word (.docx)"] = w2
            h2 = td / "s.html"
            h2.write_text("<p>Текст</p>", encoding="utf-8")
            src["HTML (.html)"] = h2
            t2 = td / "s.txt"
            t2.write_text("Строка", encoding="utf-8")
            src["TXT (.txt)"] = t2
            m2 = td / "s.md"
            m2.write_text("# Заголовок\nтекст", encoding="utf-8")
            src["Markdown (.md)"] = m2
            c2 = td / "s.csv"
            c2.write_text("a,b\n1,2\n", encoding="utf-8")
            src["CSV (.csv)"] = c2
            wb = openpyxl.Workbook()
            ws = wb.active
            ws.append(["a", "b"])
            ws.append([1, 2])
            x2 = td / "s.xlsx"
            wb.save(str(x2))
            src["Excel (.xlsx)"] = x2
            from PIL import Image
            p2 = td / "s.png"
            Image.new("RGB", (8, 8), (1, 2, 3)).save(str(p2))
            src["PNG (.png)"] = p2
            j2 = td / "s.jpg"
            Image.new("RGB", (8, 8), (1, 2, 3)).convert("RGB").save(str(j2))
            src["JPG (.jpg)"] = j2

            ext = {"Word (.docx)": ".docx", "PDF (.pdf)": ".pdf", "TXT (.txt)": ".txt",
                   "HTML (.html)": ".html", "Excel (.xlsx)": ".xlsx", "CSV (.csv)": ".csv",
                   "PNG (.png)": ".png", "JPG (.jpg)": ".jpg"}
            checked = 0
            for frm, targets in cu.CONV_SUPPORTED.items():
                for to in targets:
                    with self.subTest(frm=frm, to=to):
                        out = self._conv(frm, to, src[frm], td)
                        self.assertTrue(out.exists() and out.stat().st_size > 0,
                                        f"{frm}->{to} пустой результат")
                        checked += 1
            self.assertGreaterEqual(checked, 30)

    def test_pdf_to_word_preserves_cyrillic(self) -> None:
        with tempfile.TemporaryDirectory() as temporary_directory:
            td = Path(temporary_directory)
            pdf = td / "k.pdf"
            _make_pdf(pdf, "Привет мир")
            out = self._conv("PDF (.pdf)", "Word (.docx)", pdf, td)
            from docx import Document
            text = "\n".join(p.text for p in Document(str(out)).paragraphs)
            self.assertIn("Привет", text)

    def test_html_to_word_decodes_entities(self) -> None:
        with tempfile.TemporaryDirectory() as temporary_directory:
            td = Path(temporary_directory)
            h = td / "h.html"
            h.write_text("<p>a &amp; b &lt; c</p>", encoding="utf-8")
            out = self._conv("HTML (.html)", "Word (.docx)", h, td)
            from docx import Document
            text = "\n".join(p.text for p in Document(str(out)).paragraphs)
            self.assertIn("a & b < c", text)

    def test_unsupported_pair_raises(self) -> None:
        with tempfile.TemporaryDirectory() as temporary_directory:
            td = Path(temporary_directory)
            pdf = td / "u.pdf"
            _make_pdf(pdf)
            with self.assertRaises(RuntimeError):
                cu.convert("PDF (.pdf)", "CSV (.csv)", pdf, td / "u.out")


if __name__ == "__main__":
    unittest.main()

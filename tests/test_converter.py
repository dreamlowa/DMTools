"""Юнит-тесты для ядра converter.py: GS1 DataMatrix «Честный знак».

Покрывает чистые функции (без сети и дисковых PDF): валидацию GTIN,
разбор полного кода маркировки, проверку структуры/валидацию списков,
парсинг CSV, генерацию Zint-данных и запись CSV-файла.

Запуск: python -m unittest discover -s tests -v
"""
from __future__ import annotations

import tempfile
import unittest
from pathlib import Path

import converter as cv

# Валидный GTIN (14 цифр, корректная контрольная цифра).
GTIN = "46000000000002"


def _full_code_str(serial: str = "SER123", key91: str = "KEY",
                   sig92: str = "SIG") -> str:
    """Полная строка кода маркировки с GS-разделителями."""
    return (
        f"01{GTIN}21{serial}"
        f"{cv.GS_CHAR}91{key91}{cv.GS_CHAR}92{sig92}"
    )


def _make_marking_code(raw: str, page: int = 1) -> cv.MarkingCode:
    """Собирает MarkingCode из строкового представления кода."""
    return cv.parse_honest_sign_code(raw.encode("ascii"), page)


class GtinCheckDigitTest(unittest.TestCase):
    def test_valid_gtin(self) -> None:
        self.assertTrue(cv.gtin_check_digit_is_valid(GTIN))

    def test_invalid_gtin(self) -> None:
        self.assertFalse(cv.gtin_check_digit_is_valid("46010000000004"))

    def test_wrong_length(self) -> None:
        self.assertFalse(cv.gtin_check_digit_is_valid("4600000000000"))
        self.assertFalse(cv.gtin_check_digit_is_valid(
            "460000000000020"))

    def test_non_numeric(self) -> None:
        self.assertFalse(cv.gtin_check_digit_is_valid("46A000000000002"))


class ParseHonestSignCodeTest(unittest.TestCase):
    def test_parses_full_code(self) -> None:
        mc = _make_marking_code(_full_code_str(), page=7)
        self.assertEqual(mc.page, 7)
        self.assertEqual(mc.gtin, GTIN)
        self.assertEqual(mc.serial, "SER123")
        self.assertEqual(mc.crypto_key_91, "KEY")
        self.assertEqual(mc.crypto_code_92, "SIG")
        self.assertEqual(mc.raw.decode("ascii"), _full_code_str())

    def test_missing_ai01(self) -> None:
        with self.assertRaises(ValueError):
            cv.parse_honest_sign_code(b"02" + b"0" * 16, 1)

    def test_too_short(self) -> None:
        with self.assertRaises(ValueError):
            cv.parse_honest_sign_code(b"01", 1)

    def test_missing_ai21(self) -> None:
        raw = b"01" + GTIN.encode() + b"99X"
        with self.assertRaises(ValueError):
            cv.parse_honest_sign_code(raw, 1)

    def test_wrong_gs_split_count(self) -> None:
        # один GS вместо двух
        raw = b"01" + GTIN.encode() + b"21" + b"SER" + cv.GS + b"91K"
        with self.assertRaises(ValueError):
            cv.parse_honest_sign_code(raw, 1)

    def test_invalid_gtin_check_digit(self) -> None:
        bad = "46000000000000"
        raw = b"01" + bad.encode() + b"21SER" + cv.GS + b"91K" + cv.GS + b"92S"
        with self.assertRaises(ValueError):
            cv.parse_honest_sign_code(raw, 1)


class ValidateCodeStructureTest(unittest.TestCase):
    def test_good_code_has_no_problems(self) -> None:
        self.assertEqual(cv.validate_code_structure(_full_code_str()), [])

    def test_empty_gtin(self) -> None:
        code = f"21SER{cv.GS_CHAR}91K{cv.GS_CHAR}92S"
        self.assertTrue(any("GTIN" in p for p in cv.validate_code_structure(code)))

    def test_missing_serial(self) -> None:
        code = f"01{GTIN}{cv.GS_CHAR}91K{cv.GS_CHAR}92S"
        self.assertTrue(any("серийн" in p for p in cv.validate_code_structure(code)))

    def test_missing_91_and_92(self) -> None:
        code = f"01{GTIN}21SER"
        problems = cv.validate_code_structure(code)
        self.assertTrue(any("91" in p for p in problems))
        self.assertTrue(any("92" in p for p in problems))

    def test_invalid_gtin_check_digit_reported(self) -> None:
        code = f"0146000000000000" + f"21SER{cv.GS_CHAR}91K{cv.GS_CHAR}92S"
        self.assertTrue(any("контрольн" in p for p in cv.validate_code_structure(code)))

    def test_duplicates_detected_by_validate_codes(self) -> None:
        code = _full_code_str()
        results = cv.validate_codes([code, code])
        self.assertFalse(results[0]["is_duplicate"])
        self.assertTrue(results[1]["is_duplicate"])
        self.assertEqual(results[1]["problems"][0], "дубликат кода")

    def test_validate_codes_index(self) -> None:
        results = cv.validate_codes(["x", "y"])
        self.assertEqual([r["index"] for r in results], [1, 2])


class ParseCsvCodesTest(unittest.TestCase):
    def test_multiline(self) -> None:
        csv_text = _full_code_str() + "\n" + _full_code_str(serial="SER999")
        codes = cv.parse_csv_codes(csv_text)
        self.assertEqual(len(codes), 2)
        self.assertIn("SER999", codes[1])

    def test_strips_quotes(self) -> None:
        csv_text = '"' + _full_code_str() + '"\n'
        codes = cv.parse_csv_codes(csv_text)
        self.assertEqual(codes[0], _full_code_str())

    def test_normalizes_gs_separators(self) -> None:
        # литеральный символ GS и экранированный \\x1d должны стать GS_CHAR
        lit = _full_code_str().replace(cv.GS_CHAR, chr(29))
        escaped = _full_code_str().replace(cv.GS_CHAR, r"\x1d")
        for text in (lit, escaped):
            codes = cv.parse_csv_codes(text)
            self.assertIn(cv.GS_CHAR, codes[0])

    def test_empty_lines_ignored(self) -> None:
        csv_text = _full_code_str() + "\n\n   \n" + _full_code_str()
        self.assertEqual(len(cv.parse_csv_codes(csv_text)), 2)

    def test_empty_string(self) -> None:
        self.assertEqual(cv.parse_csv_codes(""), [])


class CodeIdentificationTest(unittest.TestCase):
    def test_short_code_drops_crypto_92(self) -> None:
        full = _full_code_str()
        short = cv.code_identification(full)
        self.assertNotIn("92", short)
        self.assertNotIn(cv.GS_CHAR + "92", short)
        self.assertIn("91" + "KEY", short)

    def test_short_code_has_serial(self) -> None:
        short = cv.code_identification(_full_code_str())
        self.assertIn("21SER123", short)

    def test_truncated_detected(self) -> None:
        # без криптохвоста 91/92 — код обрезан
        self.assertTrue(cv.is_code_truncated(f"01{GTIN}21SER"))
        # полный код с 91 — не обрезан
        self.assertFalse(cv.is_code_truncated(_full_code_str()))


class Gs1ToZintTest(unittest.TestCase):
    def test_zint_data_square_brackets(self) -> None:
        out = cv.gs1data_to_zint_data(_full_code_str())
        self.assertEqual(
            out,
            f"[01]{GTIN}[21]SER123[91]KEY[92]SIG",
        )

    def test_zint_data_empty_91_92(self) -> None:
        code = f"01{GTIN}21SER"
        out = cv.gs1data_to_zint_data(code)
        self.assertIn(f"[01]{GTIN}[21]SER", out)

    def test_zint_from_short(self) -> None:
        short = cv.code_identification(_full_code_str())
        out = cv.gs1data_to_zint_from_short(short)
        self.assertEqual(out, f"[01]{GTIN}[21]SER123[91]KEY")


class WriteCodesCsvTest(unittest.TestCase):
    def test_escape_csv_true_quoted(self) -> None:
        with tempfile.TemporaryDirectory() as td:
            target = Path(td) / "codes.csv"
            codes = [_make_marking_code(_full_code_str(), page=1)]
            cv.write_codes_csv(target, codes, escape_csv=True)
            content = target.read_text(encoding="utf-8-sig")
            # csv.writer с QUOTE_MINIMAL пишет ячейку без кавычек (нет запятых)
            self.assertIn(_full_code_str(), content)
            self.assertTrue(target.exists())

    def test_escape_csv_false_plain_sequence(self) -> None:
        codes = [_make_marking_code(_full_code_str(), page=1),
                 _make_marking_code(_full_code_str(serial="SER555"), page=2)]
        with tempfile.TemporaryDirectory() as td:
            target = Path(td) / "codes.csv"
            cv.write_codes_csv(target, codes, escape_csv=False)
            # newline="" — не перекодируем CRLF в LF (как делают universal newlines)
            content = target.read_text(encoding="utf-8-sig", newline="")
            self.assertIn(_full_code_str(), content)
            self.assertIn(_full_code_str(serial="SER555"), content)
            # строки разделены CRLF
            self.assertIn("\r\n", content)

    def test_creates_parent_dir(self) -> None:
        with tempfile.TemporaryDirectory() as td:
            target = Path(td) / "nested" / "sub" / "codes.csv"
            cv.write_codes_csv(target, [_make_marking_code(_full_code_str())],
                               escape_csv=True)
            self.assertTrue(target.exists())


if __name__ == "__main__":
    unittest.main()
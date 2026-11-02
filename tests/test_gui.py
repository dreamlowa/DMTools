from __future__ import annotations

import queue
import tempfile
import threading
import unittest
from pathlib import Path
from unittest.mock import Mock, patch

from converter import MarkingCode, PdfDecodeResult
from gui import ConverterGui, allocate_output_paths, default_combined_output_path


def marking_code(raw: bytes, page: int = 1) -> MarkingCode:
    return MarkingCode(
        page=page,
        gtin="04630196122994",
        serial="serial",
        crypto_key_91="key",
        crypto_code_92="code",
        raw=raw,
    )


class OutputPathTests(unittest.TestCase):
    def test_folder_default_is_named_after_folder_and_stored_inside_it(self) -> None:
        with tempfile.TemporaryDirectory() as temporary_directory:
            source = Path(temporary_directory) / "Orders"
            source.mkdir()

            self.assertEqual(
                default_combined_output_path(source),
                source / "Orders.csv",
            )

    def test_pdf_default_is_stored_next_to_pdf(self) -> None:
        with tempfile.TemporaryDirectory() as temporary_directory:
            source = Path(temporary_directory) / "order.pdf"
            source.touch()

            self.assertEqual(
                default_combined_output_path(source),
                source.with_suffix(".csv"),
            )

    def test_separate_outputs_stay_next_to_their_source_pdfs(self) -> None:
        with tempfile.TemporaryDirectory() as temporary_directory:
            root = Path(temporary_directory)
            first = root / "first" / "order.pdf"
            second = root / "second" / "order.pdf"
            first.parent.mkdir()
            second.parent.mkdir()
            first.touch()
            second.touch()

            self.assertEqual(
                allocate_output_paths([first, second]),
                [
                    (first, first.parent / "order_decoded_codes.csv"),
                    (second, second.parent / "order_decoded_codes.csv"),
                ],
            )


class CombinedCsvTests(unittest.TestCase):
    def test_enabling_merge_fills_empty_output_path(self) -> None:
        with tempfile.TemporaryDirectory() as temporary_directory:
            root = Path(temporary_directory)
            folder_source = root / "Orders"
            folder_source.mkdir()
            pdf_source = root / "order.pdf"
            pdf_source.touch()

            for source, expected in (
                (folder_source, folder_source / "Orders.csv"),
                (pdf_source, pdf_source.with_suffix(".csv")),
            ):
                with self.subTest(source=source):
                    app = ConverterGui.__new__(ConverterGui)
                    app.running = False
                    app.source_var = Mock(get=Mock(return_value=str(source)))
                    app.combined_output_var = Mock(
                        get=Mock(return_value=""), set=Mock()
                    )
                    app.merge_csv_var = Mock(get=Mock(return_value=True))
                    app.output_entry = Mock()
                    app.select_output_button = Mock()

                    app._update_output_controls()

                    app.combined_output_var.set.assert_called_once_with(str(expected))
                    app.output_entry.configure.assert_called_once_with(state="normal")
                    app.select_output_button.configure.assert_called_once_with(
                        state="normal"
                    )

    def test_batch_writes_codes_from_all_pdfs_to_one_csv(self) -> None:
        with tempfile.TemporaryDirectory() as temporary_directory:
            source = Path(temporary_directory) / "Orders"
            source.mkdir()
            first_pdf = source / "first.pdf"
            second_pdf = source / "second.pdf"
            first_pdf.touch()
            second_pdf.touch()
            output = source / "Orders.csv"
            first_code = marking_code(b"first")
            second_code = marking_code(b"second")

            app = ConverterGui.__new__(ConverterGui)
            app.events = queue.Queue()
            app.cancel_event = threading.Event()

            def decode(pdf_path: Path, **_kwargs: object) -> PdfDecodeResult:
                code = first_code if pdf_path == first_pdf else second_code
                return PdfDecodeResult(pdf_path, 1, (code,))

            with (
                patch("gui.decode_pdf", side_effect=decode),
                patch("gui.write_codes_csv") as write_csv,
            ):
                app._run_batch(source, output, escape_csv=False)

            write_csv.assert_called_once_with(
                output,
                [first_code, second_code],
                escape_csv=False,
            )

    def test_declining_overwrite_does_not_start_processing(self) -> None:
        with tempfile.TemporaryDirectory() as temporary_directory:
            source = Path(temporary_directory) / "Orders"
            source.mkdir()
            output = source / "Orders.csv"
            output.write_text("existing", encoding="utf-8")

            app = ConverterGui.__new__(ConverterGui)
            app.running = False
            app.source_var = Mock(get=Mock(return_value=str(source)))
            app.combined_output_var = Mock(get=Mock(return_value=str(output)))
            app.merge_csv_var = Mock(get=Mock(return_value=True))
            app.escape_csv_var = Mock(get=Mock(return_value=False))

            with (
                patch("gui.messagebox.askyesno", return_value=False) as ask,
                patch("gui.threading.Thread") as worker,
            ):
                app._start()

            ask.assert_called_once()
            worker.assert_not_called()
            self.assertEqual(output.read_text(encoding="utf-8"), "existing")

    def test_empty_combined_output_does_not_start_processing(self) -> None:
        with tempfile.TemporaryDirectory() as temporary_directory:
            source = Path(temporary_directory) / "Orders"
            source.mkdir()

            app = ConverterGui.__new__(ConverterGui)
            app.running = False
            app.source_var = Mock(get=Mock(return_value=str(source)))
            app.combined_output_var = Mock(get=Mock(return_value=""))
            app.merge_csv_var = Mock(get=Mock(return_value=True))
            app.escape_csv_var = Mock(get=Mock(return_value=False))

            with (
                patch("gui.messagebox.showerror") as show_error,
                patch("gui.threading.Thread") as worker,
            ):
                app._start()

            show_error.assert_called_once_with(
                "Ошибка", "Не указан путь к общему CSV-файлу"
            )
            worker.assert_not_called()


if __name__ == "__main__":
    unittest.main()

from __future__ import annotations

import json
import os
import queue
import sys
import threading
import calendar
from pathlib import Path
import tkinter as tk
from tkinter import filedialog, messagebox, scrolledtext, ttk

from version import APP_VERSION
from updater import check_for_update, run_self_update

from converter import (
    ConversionCancelled,
    ConversionError,
    MarkingCode,
    apply_escape_mode,
    decode_pdf,
    full_code_str,
    generate_codes_pdf,
    generate_label_pdf,
    gtin_check_digit_is_valid,
    print_pdf_windows,
    validate_codes,
    write_codes_csv,
    write_codes_xlsx,
)


def _safe_startfile(path: str) -> None:
    """Безопасно открывает файл/папку системным обработчиком."""
    try:
        os.startfile(path)  # type: ignore[attr-defined]
    except Exception:
        pass


def _make_tooltip(widget, text: str):
    """Простая подсказка при наведении курсора на виджет:
    маленькое всплывающее окно с текстом появляется с задержкой и
    скрывается при уходе курсора. Автопозиционируется ниже виджета."""
    tip = tk.Toplevel(widget)
    tip.withdraw()
    tip.overrideredirect(True)
    try:
        tip.attributes("-topmost", True)
    except Exception:
        pass
    tk.Label(
        tip, text=text, justify="left", background="#2b2b2b",
        foreground="#ffffff", relief="solid", borderwidth=1,
        padx=6, pady=3, font=("Segoe UI", 9), wraplength=360,
    ).pack()
    _after = {"id": None}

    def _cancel() -> None:
        if _after["id"] is not None:
            try:
                widget.after_cancel(_after["id"])
            except Exception:
                pass
            _after["id"] = None

    def _show() -> None:
        try:
            x = widget.winfo_rootx()
            y = widget.winfo_rooty() + widget.winfo_height() + 4
            tip.wm_geometry(f"+{x}+{y}")
            tip.deiconify()
        except Exception:
            pass

    def _schedule(_e=None) -> None:
        _cancel()
        _after["id"] = widget.after(500, _show)

    def _leave(_e=None) -> None:
        _cancel()
        tip.withdraw()

    widget.bind("<Enter>", _schedule)
    widget.bind("<Leave>", _leave)
    return tip


def _float_or_none(s: str):
    """Преобразует строку в число или None, если пусто/не число."""
    s = (s or "").strip().replace(",", ".")
    if not s:
        return None
    try:
        v = float(s)
    except ValueError:
        return None
    return v if v >= 2 else None


def find_pdf_files(source: Path) -> list[Path]:
    if source.is_file():
        return [source] if source.suffix.lower() == ".pdf" else []
    if source.is_dir():
        return sorted(
            (
                path
                for path in source.rglob("*")
                if path.is_file() and path.suffix.lower() == ".pdf"
            ),
            key=lambda path: str(path).casefold(),
        )
    return []


def allocate_output_paths(pdf_files: list[Path]) -> list[tuple[Path, Path]]:
    """Create unique names for this run; existing CSV files may be overwritten."""
    used_paths: set[str] = set()
    jobs: list[tuple[Path, Path]] = []

    for pdf_path in pdf_files:
        destination = pdf_path.parent
        base_name = f"{pdf_path.stem}_decoded_codes"
        candidate_name = f"{base_name}.csv"
        suffix = 2

        while str(destination / candidate_name).casefold() in used_paths:
            candidate_name = f"{base_name}_{suffix}.csv"
            suffix += 1

        output_path = destination / candidate_name
        used_paths.add(str(output_path).casefold())
        jobs.append((pdf_path, output_path))

    return jobs


def default_combined_output_path(source: Path) -> Path:
    if source.is_dir():
        folder_name = source.name or source.drive.rstrip(":\\/") or "result"
        return source / f"{folder_name}.csv"
    return source.with_suffix(".csv")


def config_path() -> Path:
    if getattr(sys, "frozen", False):
        base = Path(sys.executable).resolve().parent
    else:
        base = Path(__file__).resolve().parent
    return base / "PDFDataMatrixToCSV_config.json"


def log_file_path() -> Path:
    """Файл журнала операций — рядом с конфигом."""
    if getattr(sys, "frozen", False):
        base = Path(sys.executable).resolve().parent
    else:
        base = Path(__file__).resolve().parent
    return base / "PDFDataMatrixToCSV_log.txt"


def _split_drop_paths(data: str) -> list[str]:
    import re

    result: list[str] = []
    for m in re.finditer(r"\{([^}]*)\}|(\S+)", data):
        grp = m.group(1) or m.group(2)
        if grp:
            result.append(grp)
    return result


class ConverterGui:
    DATE_FORMATS = ("ММ.ГГГГ", "ММ.ГГ", "ДД.ММ.ГГГГ", "ДД.ММ.ГГ")

    def __init__(self, root: tk.Tk) -> None:
        self.root = root
        self.root.title(f"DMTools v{APP_VERSION}")
        self.root.geometry("1440x820")
        self.root.minsize(1240, 680)
        self.root.maxsize(2000, 1200)
        self._set_window_icon()

        self.source_var = tk.StringVar()
        self.combined_output_var = tk.StringVar()
        self.merge_csv_var = tk.BooleanVar(value=False)
        self.escape_mode_var = tk.StringVar(value="gs")
        self.dark_var = tk.BooleanVar(value=False)
        self.scanner_enabled = tk.BooleanVar(value=True)
        self.scanner_mode_var = tk.StringVar(value="keyboard")
        self.scanner_port_var = tk.StringVar(value="")
        self.status_var = tk.StringVar(value="Выберите PDF-файл или папку")

        self.events: queue.Queue[tuple[str, object]] = queue.Queue()
        self.cancel_event = threading.Event()
        self.worker: threading.Thread | None = None
        self.running = False
        self.last_output_dir: str | None = None
        self._last_codes: list[MarkingCode] = []
        self._multi_source_files: list[Path] = []
        self.outer: ttk.Frame | None = None
        self.dm_csv_var = tk.StringVar()
        self.dm_pdf_var = tk.StringVar()
        self.dm_label_var = tk.BooleanVar(value=True)
        self.dm_show_date_var = tk.BooleanVar(value=True)
        self.dm_show_batch_var = tk.BooleanVar(value=True)
        self.dm_show_code_var = tk.BooleanVar(value=True)
        self.dm_show_short_var = tk.BooleanVar(value=False)
        self.dm_status_var = tk.StringVar(value="Выберите CSV с кодами маркировки")
        self.dm_label_size_var = tk.StringVar(value="25 × 25")
        self.dm_group_var = tk.StringVar(value="— без группы —")
        self.dm_date_var = tk.StringVar()
        self.dm_batch_var = tk.StringVar()
        self.dm_name_var = tk.StringVar()
        self.dm_bold_name_var = tk.BooleanVar(value=False)
        self.dm_show_expiry_var = tk.BooleanVar(value=True)
        self.dm_bold_date_var = tk.BooleanVar(value=False)
        self.dm_bold_batch_var = tk.BooleanVar(value=False)
        self.dm_bold_code_var = tk.BooleanVar(value=False)
        self.dm_size_name_var = tk.StringVar(value="")
        self.dm_expiry_var = tk.StringVar()
        self.dm_bold_expiry_var = tk.BooleanVar(value=False)
        self.dm_size_expiry_var = tk.StringVar(value="")
        self.dm_size_date_var = tk.StringVar(value="")
        self.dm_size_batch_var = tk.StringVar(value="")
        self.dm_size_code_var = tk.StringVar(value="")
        self.dm_date_format_var = tk.StringVar(value="ММ.ГГГГ")
        self.dm_expiry_format_var = tk.StringVar(value="ММ.ГГГГ")
        self.dm_profile_name_var = tk.StringVar(value="")
        self.quick_code_var = tk.StringVar()
        self.quick_status_var = tk.StringVar(value="Вставьте код маркировки и нажмите «Проверить»")
        self._quick_image_ref: object | None = None
        self._profiles: dict[str, dict] = {}
        self._last_generated_pdf: str | None = None
        self._generating = False

        # видимость вкладок (галочки в меню «Вид»)
        self._tab_order = ["PDF → CSV", "CSV → DataMatrix",
                           "Быстрый DataMatrix", "Журнал"]
        self._tab_visible = {name: tk.BooleanVar(value=True)
                             for name in self._tab_order}

        self._load_config()
        self._build_menu()
        self._build_ui()
        self._apply_theme()
        self._setup_dnd()
        self._update_output_controls()
        self.root.protocol("WM_DELETE_WINDOW", self._on_close)
        self.root.after(100, self._drain_events)
        self.root.after(200, self._enable_copy_menus)
        self.root.after(1500, self._check_updates_automatic)

    def _set_window_icon(self) -> None:
        base_dir = Path(getattr(sys, "_MEIPASS", Path(__file__).resolve().parent))
        icon_path = base_dir / "centrdelfin.ico"
        if icon_path.is_file():
            try:
                self.root.iconbitmap(str(icon_path))
            except tk.TclError:
                pass

    def _build_menu(self) -> None:
        """Создаёт верхнее меню (шапку) окна: Файл, Вид, Справка с обновлениями."""
        menubar = tk.Menu(self.root)

        file_menu = tk.Menu(menubar, tearoff=0)
        file_menu.add_command(label="Открыть PDF с этикетками", command=self._open_dm_pdf)
        file_menu.add_command(label="Открыть папку с результатами", command=self._open_output_folder)
        file_menu.add_separator()
        file_menu.add_command(label="Выход", command=self._on_close)
        menubar.add_cascade(label="Файл", menu=file_menu)

        tools_menu = tk.Menu(menubar, tearoff=0)
        tools_csv = tk.Menu(tools_menu, tearoff=0)
        tools_csv.add_command(label="PDF → CSV", command=self._start)
        tools_csv.add_command(label="Остановить", command=self._stop)
        tools_csv.add_command(label="Открыть папку", command=self._open_output_folder)
        tools_csv.add_command(label="Выгрузить Excel", command=self._export_excel)
        tools_menu.add_cascade(label="PDF → CSV", menu=tools_csv)
        tools_dm = tk.Menu(tools_menu, tearoff=0)
        tools_dm.add_command(label="Сгенерировать PDF", command=self._generate_dm_pdf)
        tools_dm.add_command(label="Печать", command=self._print_dm_pdf)
        tools_dm.add_command(label="Экспорт PNG", command=self._export_dm_png)
        tools_dm.add_command(label="Открыть папку", command=self._open_dm_output_folder)
        tools_menu.add_cascade(label="CSV → этикетки", menu=tools_dm)
        menubar.add_cascade(label="Инструменты", menu=tools_menu)

        view_menu = tk.Menu(menubar, tearoff=0)
        view_menu.add_checkbutton(label="Тёмная тема", variable=self.dark_var,
                                  command=self._apply_theme)
        view_menu.add_separator()
        view_menu.add_checkbutton(
            label="PDF → CSV", variable=self._tab_visible["PDF → CSV"],
            command=lambda: self._toggle_tab("PDF → CSV"))
        view_menu.add_checkbutton(
            label="CSV → DataMatrix", variable=self._tab_visible["CSV → DataMatrix"],
            command=lambda: self._toggle_tab("CSV → DataMatrix"))
        view_menu.add_checkbutton(
            label="Быстрый DataMatrix", variable=self._tab_visible["Быстрый DataMatrix"],
            command=lambda: self._toggle_tab("Быстрый DataMatrix"))
        view_menu.add_checkbutton(
            label="Журнал", variable=self._tab_visible["Журнал"],
            command=lambda: self._toggle_tab("Журнал"))
        menubar.add_cascade(label="Вид", menu=view_menu)

        scanner_menu = tk.Menu(menubar, tearoff=0)
        scanner_menu.add_checkbutton(
            label="Приём сканера", variable=self.scanner_enabled,
            command=self._apply_scanner_enabled)
        scanner_menu.add_command(label="Настройки и проверка…",
                                 command=self._show_scanner_check)
        scanner_menu.add_separator()
        scanner_menu.add_command(label="Журнал диагностики…",
                                 command=self._show_scanner_diag)
        scanner_menu.add_command(label="Служебные штрихкоды MERTECH…",
                                 command=self._show_scanner_barcodes)
        menubar.add_cascade(label="Сканер", menu=scanner_menu)

        help_menu = tk.Menu(menubar, tearoff=0)
        help_menu.add_command(label="Проверить обновления", command=self._check_updates_manual)
        help_menu.add_separator()
        help_menu.add_command(label="О программе", command=self._show_about)
        menubar.add_cascade(label="Справка", menu=help_menu)

        self.root.config(menu=menubar)

    def _show_about(self) -> None:
        """Показывает окно «О программе»: версия и журнал изменений."""
        win = tk.Toplevel(self.root)
        win.title("О программе")
        win.resizable(False, False)
        win.transient(self.root)
        win.grab_set()

        title = f"DMTools\nверсия {APP_VERSION}"
        desc = ("Конвертация кодов маркировки GS1 DataMatrix «Честный знак»:\n"
                "PDF → CSV и CSV → этикетки с DataMatrix.")

        header = tk.Label(win, text=title, font=("Arial", 12, "bold"),
                          justify="center")
        header.pack(pady=(12, 0), padx=16)
        info = tk.Label(win, text=desc, font=("Arial", 9),
                        justify="center", fg="#555555")
        info.pack(pady=(4, 8), padx=16)

        tk.Label(win, text="Изменения по версиям:", font=("Arial", 9, "bold")
                 ).pack(anchor="w", padx=16)
        text = scrolledtext.ScrolledText(
            win, wrap="word", width=70, height=14, font=("Consolas", 9))
        text.pack(padx=16, pady=(4, 8))
        text.insert("1.0", self._changelog_text())
        text.configure(state="disabled")

        tk.Button(win, text="Закрыть", width=14,
                  command=win.destroy).pack(pady=(0, 12))

    @staticmethod
    def _changelog_text() -> str:
        """Текст журнала изменений для окна «О программе»."""
        return (
            "Текущая версия — см. «Обновления» (поддержка сканера штрихкодов\n"
            "в клавиатурном режиме: код из любого поля + Enter → проверка).\n"
            "\n"
            "v1.2.0 — верхнее меню «Справка → О программе» с версией и журналом изменений.\n"
            "\n"
            "v1.1.1 — исправлено «Не найден zint.exe» (zint упакован в EXE);\n"
            "        версия в заголовке окна и в свойствах EXE.\n"
            "\n"
            "v1.1.0 — дизайнер этикеток, проверка кодов, предпросмотр,\n"
            "        профили, печать, прогресс-бар.\n"
            "\n"
            "v1.0.0 — первая версия: PDF → CSV, CSV → этикетки.\n"
        )

    def _check_updates_manual(self) -> None:
        """Пункт меню: проверить наличие новой версии и, если есть, обновить."""
        self._run_update_check(manual=True)

    def _check_updates_automatic(self) -> None:
        """Автопроверка обновлений при старте (тихая, без всплывающих окон)."""
        self._run_update_check(manual=False)

    def _run_update_check(self, manual: bool) -> None:
        """Выполняет проверку обновлений в фоновом потоке."""
        def work() -> None:
            result = check_for_update()
            if result.get("status") == "error":
                if manual:
                    self.root.after(0, lambda: messagebox.showwarning(
                        "Обновления", result.get("message", "Не удалось проверить обновления.")))
                return
            update = result.get("update")
            if update is None:
                if manual:
                    self.root.after(0, lambda: messagebox.showinfo(
                        "Обновления", f"У вас установлена последняя версия (v{APP_VERSION})."))
                return
            self.root.after(0, lambda: self._ask_update(update))

        threading.Thread(target=work, daemon=True).start()

    def _ask_update(self, update: dict) -> None:
        """Спрашивает пользователя, обновляться ли. При согласии — запускает обновление."""
        notes = (update.get("notes") or "").strip()
        msg = (f"Доступна новая версия v{update.get('version')}.\n"
               f"Текущая версия: v{APP_VERSION}\n")
        if notes:
            msg += f"\nЧто нового:\n{notes}\n"
        msg += "\nСкачать и установить обновление сейчас?"
        if not messagebox.askyesno("Обновление программы", msg):
            return
        err = run_self_update(update["url"])
        if err:
            messagebox.showerror("Обновление", err)
        else:
            messagebox.showinfo(
                "Обновление",
                "Обновление запущено. Программа закроется, загрузит новую "
                "версию и откроется автоматически.")
            try:
                self.root.destroy()
            except Exception:
                pass

    def _enable_copy_menus(self) -> None:
        """Добавляет контекстное меню «Копировать»/«Выбрать всё» и Ctrl+C
        для всех полей ввода (Entry) в окне."""
        def add_menu(entry: tk.Entry) -> None:
            menu = tk.Menu(entry, tearoff=0)

            def do_copy():
                try:
                    entry.event_generate("<<Copy>>")
                except Exception:
                    try:
                        sel = entry.selection_get()
                        entry.clipboard_clear()
                        entry.clipboard_append(sel)
                    except Exception:
                        pass

            def do_select_all():
                entry.select_range(0, "end")
                entry.icursor("end")

            menu.add_command(label="Копировать", command=do_copy)
            menu.add_command(label="Выбрать всё", command=do_select_all)
            entry.bind("<Button-3>", lambda e, m=menu: m.tk_popup(e.x_root, e.y_root))
            entry.bind("<Control-c>", lambda e: do_copy())

        def walk(widget):
            for child in widget.winfo_children():
                if isinstance(child, ttk.Entry):
                    add_menu(child)
                walk(child)

        try:
            walk(self.root)
        except Exception:
            pass

    def _setup_dnd(self) -> None:
        try:
            self.root.drop_target_register("DND_Files")
            self.root.dnd_bind("<<Drop>>", self._on_drop)
            self.source_entry.drop_target_register("DND_Files")
            self.source_entry.dnd_bind("<<Drop>>", self._on_drop)
            if hasattr(self, "drop_zone"):
                self.drop_zone.drop_target_register("DND_Files")
                self.drop_zone.dnd_bind("<<Drop>>", self._on_drop)
                self.drop_zone.dnd_bind("<<DragEnter>>", self._on_drag_enter)
                self.drop_zone.dnd_bind("<<DragLeave>>", self._on_drag_leave)
            self.drop_enabled = True
        except Exception:
            self.drop_enabled = False

    def _on_drag_enter(self, _event: object) -> None:
        if not self.drop_enabled:
            return
        try:
            self.drop_zone.configure(fg="#2CA847", bg="#eef8ef", relief="solid", bd=2)
        except Exception:
            pass

    def _on_drag_leave(self, _event: object) -> None:
        if not self.drop_enabled:
            return
        try:
            self.drop_zone.configure(fg="#888888", bg="#fafaf9", relief="solid", bd=1)
        except Exception:
            pass

    def _on_drop(self, event: object) -> None:
        if not self.drop_enabled:
            return
        try:
            self.drop_zone.configure(fg="#888888", bg="#fafaf9", relief="solid", bd=1)
        except Exception:
            pass
        raw = getattr(event, "data", None)
        if not raw:
            return
        dropped = [Path(p) for p in _split_drop_paths(str(raw))]
        pdfs = [p for p in dropped if p.suffix.lower() == ".pdf"]
        if not pdfs:
            return
        self._multi_source_files = pdfs
        if len(pdfs) == 1:
            self.source_var.set(str(pdfs[0]))
            self.combined_output_var.set(str(default_combined_output_path(pdfs[0])))
        else:
            self.source_var.set(
                f"{len(pdfs)} файлов: "
                + "; ".join(str(p) for p in pdfs[:3])
                + (" ..." if len(pdfs) > 3 else "")
            )
            self.combined_output_var.set(
                str(default_combined_output_path(pdfs[0].parent))
            )
        self._update_output_controls()

    def _apply_theme(self) -> None:
        dark = self.dark_var.get()
        style = ttk.Style(self.root)

        if dark:
            bg = "#1f1f1f"
            fg = "#e8e8e8"
            field = "#2b2b2b"
            sel = "#333333"
            border = "#3a3a3a"
            try:
                style.theme_use("clam")
            except tk.TclError:
                pass
            style.configure(".", background=bg, foreground=fg)
            style.configure("TFrame", background=bg)
            style.configure("TLabelframe", background=bg, foreground=fg)
            style.configure("TLabelframe.Label", background=bg, foreground=fg)
            style.configure("TLabel", background=bg, foreground=fg)
            style.configure(
                "TCheckbutton", background=bg, foreground=fg,
                focuscolor=bg, indicatorrelief="flat", indicatormargin=4,
            )
            style.map(
                "TCheckbutton",
                background=[("active", bg), ("pressed", bg)],
                indicatorcolor=[("selected", "#2CA847"), ("active", field), ("!selected", field)],
            )
            style.configure(
                "TRadiobutton", background=bg, foreground=fg,
                focuscolor=bg, indicatorrelief="flat", indicatormargin=4,
            )
            style.map(
                "TRadiobutton",
                background=[("active", bg), ("pressed", bg)],
                indicatorcolor=[("selected", "#2CA847"), ("active", field), ("!selected", field)],
            )
            style.configure(
                "TButton", background=sel, foreground=fg, bordercolor=border,
                focuscolor=bg, lightcolor=bg, darkcolor=bg, relief="flat",
                padding=(8, 3),
            )
            style.map("TButton", background=[("active", "#444444"), ("pressed", "#2a2a2a")])
            # акцентная (зелёная) кнопка главного действия
            style.configure(
                "Accent.TButton", background="#2CA847", foreground="#ffffff",
                bordercolor="#2CA847", focuscolor=bg, relief="flat",
                lightcolor="#2CA847", darkcolor="#2CA847", padding=(12, 4),
                font=("Segoe UI", 9, "bold"),
            )
            style.map(
                "Accent.TButton",
                background=[("active", "#24913c"), ("pressed", "#1f7e34")],
                foreground=[("active", "#ffffff"), ("pressed", "#ffffff")],
            )
            # синяя — действие с дополнительной настройкой (инструменты)
            style.configure(
                "Info.TButton", background="#2b6cb0", foreground="#ffffff",
                bordercolor="#2b6cb0", focuscolor=bg, relief="flat",
                lightcolor="#2b6cb0", darkcolor="#2b6cb0", padding=(10, 4),
                font=("Segoe UI", 9, "bold"),
            )
            style.map(
                "Info.TButton",
                background=[("active", "#245a94"), ("pressed", "#1d4b7a")],
                foreground=[("active", "#ffffff"), ("pressed", "#ffffff")],
            )
            # оранжевая — предупреждающее действие
            style.configure(
                "Warn.TButton", background="#d97706", foreground="#ffffff",
                bordercolor="#d97706", focuscolor=bg, relief="flat",
                lightcolor="#d97706", darkcolor="#d97706", padding=(10, 4),
                font=("Segoe UI", 9, "bold"),
            )
            style.map(
                "Warn.TButton",
                background=[("active", "#b75f00"), ("pressed", "#995200")],
                foreground=[("active", "#ffffff"), ("pressed", "#ffffff")],
            )
            # красная — отмена / опасное действие / стоп
            style.configure(
                "Danger.TButton", background="#c0392b", foreground="#ffffff",
                bordercolor="#c0392b", focuscolor=bg, relief="flat",
                lightcolor="#c0392b", darkcolor="#c0392b", padding=(10, 4),
                font=("Segoe UI", 9, "bold"),
            )
            style.map(
                "Danger.TButton",
                background=[("active", "#a93226"), ("pressed", "#922b21")],
                foreground=[("active", "#ffffff"), ("pressed", "#ffffff")],
            )
            # нейтральная — второстепенное действие
            style.configure(
                "Neutral.TButton", background="#6c757d", foreground="#ffffff",
                bordercolor="#6c757d", focuscolor=bg, relief="flat",
                lightcolor="#6c757d", darkcolor="#6c757d", padding=(8, 3),
            )
            style.map(
                "Neutral.TButton",
                background=[("active", "#5a6268"), ("pressed", "#4c5258")],
                foreground=[("active", "#ffffff"), ("pressed", "#ffffff")],
            )
            # подпись ошибки валидации поля
            style.configure("Error.TLabel", background=bg, foreground="#ff8080")
            style.configure(
                "TEntry", fieldbackground=field, foreground=fg,
                bordercolor=border, insertcolor=fg, padding=(4, 3),
            )
            style.map("TEntry", fieldbackground=[("focus", field)], bordercolor=[("focus", "#2CA847")])
            style.configure("TProgressbar", background="#2CA847", troughcolor=field)
            style.configure("Vertical.TScrollbar", background=sel, troughcolor=bg, arrowcolor=fg, bordercolor=bg)
            style.configure("Horizontal.TScrollbar", background=sel, troughcolor=bg, arrowcolor=fg, bordercolor=bg)
            # вкладки Notebook — читаемые в тёмной теме
            style.configure(
                "TNotebook", background=bg, borderwidth=0, tabmargins=(2, 2, 2, 0),
            )
            style.configure(
                "TNotebook.Tab",
                background="#2a2a2a", foreground="#cfcfcf", padding=(14, 4),
                borderwidth=0, focuscolor=bg, bordercolor=border,
            )
            style.map(
                "TNotebook.Tab",
                background=[("selected", "#3a3a3a"), ("active", "#333333")],
                foreground=[("selected", "#ffffff"), ("active", "#ffffff")],
                padding=[("selected", (16, 6)), ("active", (14, 4))],
            )
        else:
            bg = "#f0f0f0"
            fg = "#000000"
            field = "#ffffff"
            try:
                style.theme_use("vista")
            except tk.TclError:
                pass
            style.configure("TFrame", background=bg)
            style.configure("Labelframe", background=bg, foreground=fg)
            style.configure("Labelframe.Label", background=bg, foreground=fg)
            style.configure("TLabel", background=bg, foreground=fg)
            style.configure("TCheckbutton", background=bg, foreground=fg)
            style.configure("TEntry", fieldbackground=field, foreground=fg, insertcolor=fg)
            style.configure("TProgressbar", background="#2CA847", troughcolor=field)
            # акцентная (зелёная) кнопка главного действия
            style.configure(
                "Accent.TButton", background="#2CA847", foreground="#ffffff",
                bordercolor="#2CA847", focuscolor=bg, relief="flat",
                lightcolor="#2CA847", darkcolor="#2CA847", padding=(12, 4),
                font=("Segoe UI", 9, "bold"),
            )
            style.map(
                "Accent.TButton",
                background=[("active", "#24913c"), ("pressed", "#1f7e34")],
                foreground=[("active", "#ffffff"), ("pressed", "#ffffff")],
            )
            # синяя — действие с дополнительной настройкой (инструменты)
            style.configure(
                "Info.TButton", background="#2b6cb0", foreground="#ffffff",
                bordercolor="#2b6cb0", focuscolor=bg, relief="flat",
                lightcolor="#2b6cb0", darkcolor="#2b6cb0", padding=(10, 4),
                font=("Segoe UI", 9, "bold"),
            )
            style.map(
                "Info.TButton",
                background=[("active", "#245a94"), ("pressed", "#1d4b7a")],
                foreground=[("active", "#ffffff"), ("pressed", "#ffffff")],
            )
            # оранжевая — предупреждающее действие
            style.configure(
                "Warn.TButton", background="#d97706", foreground="#ffffff",
                bordercolor="#d97706", focuscolor=bg, relief="flat",
                lightcolor="#d97706", darkcolor="#d97706", padding=(10, 4),
                font=("Segoe UI", 9, "bold"),
            )
            style.map(
                "Warn.TButton",
                background=[("active", "#b75f00"), ("pressed", "#995200")],
                foreground=[("active", "#ffffff"), ("pressed", "#ffffff")],
            )
            # красная — отмена / опасное действие / стоп
            style.configure(
                "Danger.TButton", background="#c0392b", foreground="#ffffff",
                bordercolor="#c0392b", focuscolor=bg, relief="flat",
                lightcolor="#c0392b", darkcolor="#c0392b", padding=(10, 4),
                font=("Segoe UI", 9, "bold"),
            )
            style.map(
                "Danger.TButton",
                background=[("active", "#a93226"), ("pressed", "#922b21")],
                foreground=[("active", "#ffffff"), ("pressed", "#ffffff")],
            )
            # нейтральная — второстепенное действие
            style.configure(
                "Neutral.TButton", background="#6c757d", foreground="#ffffff",
                bordercolor="#6c757d", focuscolor=bg, relief="flat",
                lightcolor="#6c757d", darkcolor="#6c757d", padding=(8, 3),
            )
            style.map(
                "Neutral.TButton",
                background=[("active", "#5a6268"), ("pressed", "#4c5258")],
                foreground=[("active", "#ffffff"), ("pressed", "#ffffff")],
            )
            # подпись ошибки валидации поля
            style.configure("Error.TLabel", background=bg, foreground="#c0392b")
            # вкладки — активная крупнее остальных
            style.configure(
                "TNotebook.Tab",
                background="#e4e4e4", foreground="#333333", padding=(14, 4),
                borderwidth=1, focuscolor=bg, bordercolor="#c0c0c0",
            )
            style.map(
                "TNotebook.Tab",
                background=[("selected", "#ffffff"), ("active", "#eeeeee")],
                foreground=[("selected", "#000000"), ("active", "#000000")],
                padding=[("selected", (16, 6)), ("active", (14, 4))],
                bordercolor=[("selected", "#2CA847")],
            )

        self.root.configure(bg=bg)
        if self.outer is not None:
            self.outer.configure(style="TFrame")
        if hasattr(self, "global_log"):
            try:
                self.global_log.configure(bg=field, fg=fg, insertbackground=fg)
            except Exception:
                pass
        if hasattr(self, "dm_preview_canvas"):
            try:
                self.dm_preview_canvas.configure(bg=field)
            except Exception:
                pass
        if hasattr(self, "dm_full_preview_canvas"):
            try:
                self.dm_full_preview_canvas.configure(bg=field)
            except Exception:
                pass
        if hasattr(self, "drop_zone"):
            try:
                self.drop_zone.configure(bg=field, fg=fg)
            except Exception:
                pass
        if hasattr(self, "_header_title"):
            try:
                self._header_title.configure(bg=bg, fg=fg)
            except Exception:
                pass
        if hasattr(self, "_header_logo"):
            try:
                self._header_logo.configure(bg=bg)
            except Exception:
                pass

        for widget in self.root.winfo_children():
            self._refresh_style(widget)

    def _refresh_style(self, widget: tk.Misc) -> None:
        try:
            if isinstance(widget, ttk.Label):
                cur = str(widget.cget("style"))
                if cur == "Error.TLabel":
                    widget.configure(style="Error.TLabel")
                else:
                    widget.configure(style="TLabel")
            elif isinstance(widget, ttk.Entry):
                widget.configure(style="TEntry")
            elif isinstance(widget, ttk.Checkbutton):
                widget.configure(style="TCheckbutton")
            elif isinstance(widget, ttk.Radiobutton):
                widget.configure(style="TRadiobutton")
            elif isinstance(widget, ttk.Button):
                if str(widget.cget("style")) != "Accent.TButton":
                    widget.configure(style="TButton")
            elif isinstance(widget, ttk.Frame):
                widget.configure(style="TFrame")
        except tk.TclError:
            pass
        for child in widget.winfo_children():
            self._refresh_style(child)

    def _build_header(self, parent: ttk.Frame) -> None:
        """Компактная верхняя панель с логотипом и названием «DMTools»."""
        header = ttk.Frame(parent, padding=(8, 4))
        header.grid(row=0, column=0, sticky="ew")
        header.columnconfigure(1, weight=1)
        self._header_photo = None
        # цвет фона шапки берём из текущей темы, чтобы label не имел пустого bg
        try:
            style = ttk.Style(self.root)
            hdr_bg = style.lookup("TFrame", "background") or "#f0f0f0"
            hdr_fg = style.lookup("TLabel", "foreground") or "#000000"
        except Exception:
            hdr_bg, hdr_fg = "#f0f0f0", "#000000"
        try:
            base_dir = Path(getattr(sys, "_MEIPASS", Path(__file__).resolve().parent))
            icon_path = base_dir / "centrdelfin.ico"
            if icon_path.is_file():
                from PIL import Image, ImageTk
                img = Image.open(str(icon_path)).convert("RGBA")
                img.thumbnail((28, 28))
                self._header_photo = ImageTk.PhotoImage(img)
                self._header_logo = tk.Label(
                    header, image=self._header_photo, bg=hdr_bg,
                    relief="flat",
                )
                self._header_logo.grid(row=0, column=0, padx=(2, 8))
        except Exception:
            pass
        self._header_title = tk.Label(
            header, text="DMTools", font=("Segoe UI", 13, "bold"),
            bg=hdr_bg, fg=hdr_fg,
        )
        self._header_title.grid(row=0, column=1, sticky="w")

    def _build_ui(self) -> None:
        root_container = ttk.Frame(self.root, padding=8)
        root_container.grid(row=0, column=0, sticky="nsew")
        self.root.columnconfigure(0, weight=1)
        self.root.rowconfigure(0, weight=1)

        # Верхняя панель с заголовком и логотипом
        self._build_header(root_container)
        self.dark_check = None

        self.notebook = ttk.Notebook(root_container)
        self.notebook.grid(row=1, column=0, sticky="nsew")
        root_container.rowconfigure(1, weight=1)

        # --- Вкладка 1: PDF -> CSV ---
        tab1 = ttk.Frame(self.notebook, padding=10)
        self.notebook.add(tab1, text="PDF → CSV")

        outer = ttk.Frame(tab1)
        self.outer = outer
        outer.grid(row=0, column=0, sticky="nsew")
        outer.columnconfigure(1, weight=1)
        outer.rowconfigure(10, weight=1)

        # Drop zone banner
        drop_outer = ttk.Frame(outer, style="TFrame")
        drop_outer.grid(row=0, column=0, columnspan=3, sticky="ew", pady=(0, 12))
        self.drop_zone = tk.Label(
            drop_outer,
            text="⬇  Перетащите PDF-файлы сюда  ⬇\nили нажмите «Выбрать PDF» ниже",
            font=("Segoe UI", 13, "bold"),
            justify="center",
            fg="#888888",
            bg="#fafaf9",
            relief="solid",
            bd=1,
            height=3,
            padx=20,
            pady=14,
            cursor="hand2",
        )
        self.drop_zone.pack(fill="x")

        ttk.Label(outer, text="Источник:").grid(
            row=1, column=0, sticky="e", padx=(0, 8), pady=(0, 8)
        )
        self.source_entry = ttk.Entry(outer, textvariable=self.source_var)
        self.source_entry.grid(row=1, column=1, sticky="ew", pady=(0, 8))

        source_buttons = ttk.Frame(outer)
        source_buttons.grid(row=1, column=2, padx=(8, 0), pady=(0, 8))
        self.select_file_button = ttk.Button(
            source_buttons, text="Выбрать PDF", command=self._select_pdf,
            style="Neutral.TButton"
        )
        self.select_file_button.grid(row=0, column=0, padx=(0, 6))
        self.select_folder_button = ttk.Button(
            source_buttons, text="Выбрать папку", command=self._select_source_folder,
            style="Neutral.TButton"
        )
        self.select_folder_button.grid(row=0, column=1)

        self.merge_csv_check = ttk.Checkbutton(
            outer,
            text="Объединять в один CSV файл",
            variable=self.merge_csv_var,
            command=self._update_output_controls,
        )
        self.merge_csv_check.grid(
            row=2, column=1, columnspan=2, sticky="w", pady=(0, 8)
        )

        ttk.Label(outer, text="Общий CSV:").grid(
            row=3, column=0, sticky="e", padx=(0, 8), pady=(0, 8)
        )
        self.output_entry = ttk.Entry(outer, textvariable=self.combined_output_var)
        self.output_entry.grid(row=3, column=1, sticky="ew", pady=(0, 8))
        self.select_output_button = ttk.Button(
            outer, text="Выбрать файл", command=self._select_output_file,
            style="Neutral.TButton"
        )
        self.select_output_button.grid(
            row=3, column=2, sticky="ew", padx=(8, 0), pady=(0, 8)
        )

        escape_group = ttk.LabelFrame(outer, text="Разделитель GS в CSV")
        escape_group.grid(row=4, column=1, columnspan=2, sticky="w", pady=(0, 8))
        self.escape_gs_radio = ttk.Radiobutton(
            escape_group, text="Символ GS (обычный)", value="gs",
            variable=self.escape_mode_var,
        )
        self.escape_gs_radio.grid(row=0, column=0, sticky="w", padx=8, pady=2)
        self.escape_1c_radio = ttk.Radiobutton(
            escape_group, text="Для 1С (GS заменён на \\x1d)", value="1c",
            variable=self.escape_mode_var,
        )
        self.escape_1c_radio.grid(row=1, column=0, sticky="w", padx=8, pady=2)
        self.escape_ozon_radio = ttk.Radiobutton(
            escape_group, text="Для OZON (GS заменён на \\u001d)", value="ozon",
            variable=self.escape_mode_var,
        )
        self.escape_ozon_radio.grid(row=2, column=0, sticky="w", padx=8, pady=2)

        ttk.Label(
            outer,
            text=(
                "Для папки PDF-файлы ищутся рекурсивно. "
                "По умолчанию CSV сохраняются рядом с исходными PDF."
            ),
        ).grid(row=7, column=0, columnspan=3, sticky="w", pady=(0, 10))

        actions = ttk.Frame(outer)
        actions.grid(row=8, column=0, columnspan=3, sticky="ew", pady=(0, 8))
        actions.columnconfigure(2, weight=1)
        self.start_button = ttk.Button(
            actions, text="Начать конвертацию", command=self._start,
            style="Accent.TButton",
        )
        self.start_button.grid(row=0, column=0, padx=(0, 8))
        self.stop_button = ttk.Button(
            actions, text="Остановить", command=self._stop, state="disabled",
            style="Danger.TButton",
        )
        self.stop_button.grid(row=0, column=1)
        self.open_folder_button = ttk.Button(
            actions, text="Перейти в папку", command=self._open_output_folder,
            style="Neutral.TButton",
        )
        self.open_folder_button.grid(row=0, column=2, sticky="w", padx=(12, 0))
        self.open_codes_button = ttk.Button(
            actions, text="Открыть код", command=self._open_codes_dialog,
            style="Info.TButton",
        )
        self.open_codes_button.grid(row=0, column=3, sticky="w", padx=(12, 0))
        self.excel_button = ttk.Button(
            actions, text="Выгрузить Excel", command=self._export_excel,
            style="Warn.TButton",
        )
        self.excel_button.grid(row=0, column=4, sticky="w", padx=(12, 0))
        ttk.Label(actions, textvariable=self.status_var).grid(
            row=0, column=5, sticky="e", padx=(12, 0)
        )

        self.progress = ttk.Progressbar(outer, mode="determinate")
        self.progress.grid(row=9, column=0, columnspan=3, sticky="ew", pady=(0, 10))

        # --- Вкладка 2: CSV -> DataMatrix ---
        tab2 = ttk.Frame(self.notebook, padding=12)
        self.notebook.add(tab2, text="CSV → DataMatrix")
        tab2.columnconfigure(0, weight=1)

        # Внутренняя панель: слева — ввод, справа — предпросмотр
        tab2_body = ttk.Frame(tab2)
        tab2_body.grid(row=0, column=0, sticky="nsew")
        tab2.rowconfigure(0, weight=1)
        tab2_body.columnconfigure(0, weight=1)
        tab2_body.columnconfigure(1, weight=0)

        # ===== Левая часть: настройки =====
        # Единая сетка формы: колонка 0 — подписи (прижаты к правому краю,
        # чтобы все поля начинались на одной вертикальной линии), колонка 1 —
        # поля ввода, колонки 2-4 — вспомогательные элементы. Доп. элементы у
        # полей (формат даты, «Жирн.», размер шрифта) собраны в компактный
        # под-фрейм прямо у поля, чтобы ничего не «вываливалось» в дальние колонки.
        left = ttk.Frame(tab2_body, padding=(0, 0, 12, 0))
        left.grid(row=0, column=0, sticky="nsew")
        left.columnconfigure(0, weight=1)

        # ===== Блок 1: Данные этикетки =====
        # Единая сетка формы внутри рамки: колонка 0 — подписи (прижаты к правому
        # краю, чтобы все поля начинались на одной вертикальной линии), колонка 1 —
        # поля ввода, колонки 2-4 — вспомогательные элементы. Доп. элементы у
        # полей (формат даты, «Жирн.», размер шрифта) собраны в компактный
        # под-фрейм прямо у поля, чтобы ничего не «вываливалось» в дальние колонки.
        data_frame = ttk.LabelFrame(left, text="Данные этикетки", padding=(10, 6))
        data_frame.grid(row=0, column=0, sticky="ew", pady=(0, 10))
        data_frame.columnconfigure(0, weight=0)
        data_frame.columnconfigure(1, weight=1)
        for _c in (2, 3, 4):
            data_frame.columnconfigure(_c, weight=0)

        # --- Строка: файл CSV ---
        ttk.Label(data_frame, text="Файл CSV (по одному коду\nв строке):").grid(
            row=0, column=0, sticky="e", padx=(0, 8), pady=(0, 8)
        )
        self.dm_csv_entry = ttk.Entry(data_frame, textvariable=self.dm_csv_var)
        self.dm_csv_entry.grid(row=0, column=1, sticky="ew", pady=(0, 8))
        self.dm_pick_files_btn = ttk.Button(
            data_frame, text="Выбрать файлы", command=self._select_dm_csv,
            style="Neutral.TButton"
        )
        self.dm_pick_files_btn.grid(row=0, column=2, padx=(8, 0), pady=(0, 8))
        self.dm_pick_folder_btn = ttk.Button(
            data_frame, text="Выбрать папку", command=self._select_dm_folder,
            style="Neutral.TButton"
        )
        self.dm_pick_folder_btn.grid(row=0, column=3, padx=(8, 0), pady=(0, 8))

        # --- Строка: группа товара ---
        ttk.Label(data_frame, text="Группа товара:").grid(
            row=1, column=0, sticky="e", padx=(0, 8), pady=(0, 8)
        )
        groups = [
            "— без группы —", "Медизделия", "Духи/парфюмерия", "Косметика",
            "Лёгкая промышленность", "Антисептики",
        ]
        self.dm_group_combo = ttk.Combobox(
            data_frame, textvariable=self.dm_group_var, values=groups,
            state="readonly", width=16,
        )
        self.dm_group_combo.grid(row=1, column=1, sticky="w", pady=(0, 8))
        self.dm_group_combo.bind("<<ComboboxSelected>>", self._on_dm_group_change)

        # --- Строка: размер этикетки ---
        ttk.Label(data_frame, text="Размер этикетки (мм):").grid(
            row=2, column=0, sticky="e", padx=(0, 8), pady=(0, 8)
        )
        self.dm_size_combo = ttk.Combobox(
            data_frame, textvariable=self.dm_label_size_var,
            values=self._sizes_for_group(self.dm_group_var.get()),
            state="readonly", width=16,
        )
        self.dm_size_combo.grid(row=2, column=1, sticky="w", pady=(0, 8))
        self.dm_size_combo.bind("<<ComboboxSelected>>", self._on_preview_input_change)

        # --- Строка: дата производства (поле + формат + «Жирн.» + размер) ---
        ttk.Label(data_frame, text="Дата произв.:").grid(
            row=3, column=0, sticky="e", padx=(0, 8), pady=(0, 8)
        )
        date_row = ttk.Frame(data_frame)
        date_row.grid(row=3, column=1, columnspan=2, sticky="w", pady=(0, 8))
        date_sub = ttk.Frame(date_row)
        date_sub.pack(anchor="w")
        self.dm_date_entry = ttk.Entry(date_sub, textvariable=self.dm_date_var, width=11)
        self.dm_date_entry.pack(side="left")
        self.dm_date_entry.bind("<KeyRelease>", self._on_date_format)
        self.dm_date_fmt_combo = ttk.Combobox(
            date_sub, textvariable=self.dm_date_format_var,
            values=self.DATE_FORMATS, state="readonly", width=9,
        )
        self.dm_date_fmt_combo.pack(side="left", padx=(8, 0))
        self.dm_date_fmt_combo.bind("<<ComboboxSelected>>", self._on_date_format)
        self.dm_bold_date_check = ttk.Checkbutton(
            date_sub, text="Жирн.", variable=self.dm_bold_date_var, command=self._on_show_change
        )
        self.dm_bold_date_check.pack(side="left", padx=(12, 0))
        self.dm_size_date_spin = ttk.Spinbox(
            date_sub, from_=2, to=20, increment=0.5, width=4,
            textvariable=self.dm_size_date_var, command=self._on_show_change,
        )
        self.dm_size_date_spin.pack(side="left", padx=(8, 0))
        self.dm_date_error_label = ttk.Label(
            date_row, text="", style="Error.TLabel", anchor="w", wraplength=360,
        )
        self.dm_date_error_label.pack(anchor="w", pady=(2, 0))
        self.dm_date_tooltip = _make_tooltip(
            self.dm_date_entry,
            "Формат по выбору: ММ.ГГГГ / ММ.ГГ / ДД.ММ.ГГГГ / ДД.ММ.ГГ.\n"
            "Дата производства — не позже сегодня.",
        )

        # --- Строка: срок годности (аналогично дате производства) ---
        ttk.Label(data_frame, text="Срок годности:").grid(
            row=4, column=0, sticky="e", padx=(0, 8), pady=(0, 8)
        )
        expiry_row = ttk.Frame(data_frame)
        expiry_row.grid(row=4, column=1, columnspan=2, sticky="w", pady=(0, 8))
        expiry_sub = ttk.Frame(expiry_row)
        expiry_sub.pack(anchor="w")
        self.dm_expiry_entry = ttk.Entry(
            expiry_sub, textvariable=self.dm_expiry_var, width=11
        )
        self.dm_expiry_entry.pack(side="left")
        self.dm_expiry_entry.bind("<KeyRelease>", self._on_date_format)
        self.dm_expiry_fmt_combo = ttk.Combobox(
            expiry_sub, textvariable=self.dm_expiry_format_var,
            values=self.DATE_FORMATS, state="readonly", width=9,
        )
        self.dm_expiry_fmt_combo.pack(side="left", padx=(8, 0))
        self.dm_expiry_fmt_combo.bind("<<ComboboxSelected>>", self._on_date_format)
        self.dm_bold_expiry_check = ttk.Checkbutton(
            expiry_sub, text="Жирн.", variable=self.dm_bold_expiry_var, command=self._on_show_change
        )
        self.dm_bold_expiry_check.pack(side="left", padx=(12, 0))
        self.dm_size_expiry_spin = ttk.Spinbox(
            expiry_sub, from_=2, to=20, increment=0.5, width=4,
            textvariable=self.dm_size_expiry_var, command=self._on_show_change,
        )
        self.dm_size_expiry_spin.pack(side="left", padx=(8, 0))
        self.dm_expiry_error_label = ttk.Label(
            expiry_row, text="", style="Error.TLabel", anchor="w", wraplength=360,
        )
        self.dm_expiry_error_label.pack(anchor="w", pady=(2, 0))
        self.dm_expiry_tooltip = _make_tooltip(
            self.dm_expiry_entry,
            "Формат по выбору: ММ.ГГГГ / ММ.ГГ / ДД.ММ.ГГГГ / ДД.ММ.ГГ.\n"
            "Срок годности может быть любым.",
        )

        # --- Строка: партия ---
        ttk.Label(data_frame, text="Партия:").grid(
            row=5, column=0, sticky="e", padx=(0, 8), pady=(0, 8)
        )
        batch_row = ttk.Frame(data_frame)
        batch_row.grid(row=5, column=1, sticky="w", pady=(0, 8))
        self.dm_batch_entry = ttk.Entry(batch_row, textvariable=self.dm_batch_var, width=28)
        self.dm_batch_entry.pack(side="left")
        self.dm_batch_entry.bind("<KeyRelease>", self._on_preview_input_change)
        self.dm_bold_batch_check = ttk.Checkbutton(
            batch_row, text="Жирн.", variable=self.dm_bold_batch_var, command=self._on_show_change
        )
        self.dm_bold_batch_check.pack(side="left", padx=(12, 0))
        self.dm_size_batch_spin = ttk.Spinbox(
            batch_row, from_=2, to=20, increment=0.5, width=4,
            textvariable=self.dm_size_batch_var, command=self._on_show_change,
        )
        self.dm_size_batch_spin.pack(side="left", padx=(8, 0))

        # --- Строка: наименование ---
        ttk.Label(data_frame, text="Наименование:").grid(
            row=6, column=0, sticky="e", padx=(0, 8), pady=(0, 8)
        )
        name_row = ttk.Frame(data_frame)
        name_row.grid(row=6, column=1, sticky="w", pady=(0, 8))
        self.dm_name_entry = ttk.Entry(name_row, textvariable=self.dm_name_var, width=28)
        self.dm_name_entry.pack(side="left")
        self.dm_name_entry.bind("<KeyRelease>", self._on_preview_input_change)
        self.dm_bold_name_check = ttk.Checkbutton(
            name_row, text="Жирн.", variable=self.dm_bold_name_var, command=self._on_show_change
        )
        self.dm_bold_name_check.pack(side="left", padx=(12, 0))
        self.dm_size_name_spin = ttk.Spinbox(
            name_row, from_=2, to=20, increment=0.5, width=4,
            textvariable=self.dm_size_name_var, command=self._on_show_change,
        )
        self.dm_size_name_spin.pack(side="left", padx=(8, 0))

        # --- Строка: ручной ввод кода маркировки ---
        ttk.Label(data_frame, text="Код маркировки:").grid(
            row=7, column=0, sticky="e", padx=(0, 8), pady=(0, 8)
        )
        manual_row = ttk.Frame(data_frame)
        manual_row.grid(row=7, column=1, columnspan=3, sticky="ew", pady=(0, 8))
        manual_row.columnconfigure(0, weight=1)
        self.dm_manual_code_var = tk.StringVar()
        self.dm_manual_code_entry = ttk.Entry(
            manual_row, textvariable=self.dm_manual_code_var
        )
        self.dm_manual_code_entry.grid(row=0, column=0, sticky="ew", padx=(0, 8))
        self.dm_manual_code_entry.bind(
            "<Return>", lambda _e: self._manual_code_generate()
        )
        ttk.Button(
            manual_row, text="Сгенерировать DataMatrix",
            command=self._manual_code_generate, style="Accent.TButton"
        ).grid(row=0, column=1)

        # ===== Блок 2: Вывод =====
        output_frame = ttk.LabelFrame(left, text="Вывод", padding=(10, 6))
        output_frame.grid(row=1, column=0, sticky="ew")
        output_frame.columnconfigure(0, weight=0)
        output_frame.columnconfigure(1, weight=1)
        for _c in (2, 3, 4):
            output_frame.columnconfigure(_c, weight=0)

        # --- Строка: сохранить PDF ---
        ttk.Label(output_frame, text="Сохранить PDF:").grid(
            row=0, column=0, sticky="e", padx=(0, 8), pady=(0, 8)
        )
        self.dm_pdf_entry = ttk.Entry(output_frame, textvariable=self.dm_pdf_var)
        self.dm_pdf_entry.grid(row=0, column=1, sticky="ew", pady=(0, 8))
        ttk.Button(output_frame, text="Выбрать PDF", command=self._select_dm_pdf,
                  style="Neutral.TButton").grid(
            row=0, column=2, padx=(8, 0), pady=(0, 8)
        )

        # --- Строка: показывать на этикетке (галочки в 2 аккуратных ряда) ---
        ttk.Label(output_frame, text="Показывать\nна этикетке:").grid(
            row=1, column=0, sticky="e", padx=(0, 8), pady=(0, 8)
        )
        show_opts = ttk.Frame(output_frame)
        show_opts.grid(row=1, column=1, columnspan=4, sticky="w", pady=(0, 8))
        show_opts.columnconfigure(3, weight=1)
        self.dm_show_date_check = ttk.Checkbutton(
            show_opts, text="Дата произв.", variable=self.dm_show_date_var, command=self._on_show_change
        )
        self.dm_show_date_check.grid(row=0, column=0, sticky="w", padx=(0, 16))
        self.dm_show_expiry_check = ttk.Checkbutton(
            show_opts, text="Срок годн.", variable=self.dm_show_expiry_var, command=self._on_show_change
        )
        self.dm_show_expiry_check.grid(row=0, column=1, sticky="w", padx=(0, 16))
        self.dm_show_batch_check = ttk.Checkbutton(
            show_opts, text="Партия", variable=self.dm_show_batch_var, command=self._on_show_change
        )
        self.dm_show_batch_check.grid(row=0, column=2, sticky="w", padx=(0, 16))
        self.dm_show_short_check = ttk.Checkbutton(
            show_opts, text="Укороченный код", variable=self.dm_show_short_var,
            command=self._on_show_change,
        )
        self.dm_show_short_check.grid(row=0, column=3, sticky="w", pady=(4, 0))
        self.dm_show_code_check = ttk.Checkbutton(
            show_opts, text="Код маркировки", variable=self.dm_show_code_var, command=self._on_show_change
        )
        self.dm_show_code_check.grid(row=1, column=0, sticky="w", pady=(6, 0), padx=(0, 4))
        self.dm_bold_code_check = ttk.Checkbutton(
            show_opts, text="жирн.", variable=self.dm_bold_code_var, command=self._on_show_change
        )
        self.dm_bold_code_check.grid(row=1, column=1, sticky="w", pady=(6, 0), padx=(0, 4))
        self.dm_size_code_spin = ttk.Spinbox(
            show_opts, from_=2, to=20, increment=0.5, width=4,
            textvariable=self.dm_size_code_var, command=self._on_show_change,
        )
        self.dm_size_code_spin.grid(row=1, column=2, sticky="w", pady=(6, 0))

        # --- Строка: кнопки действий (в одну ровную линейку с равными отступами) ---
        dm_actions = ttk.Frame(output_frame)
        dm_actions.grid(row=2, column=0, columnspan=5, sticky="ew", pady=(4, 8))
        self.dm_generate_button = ttk.Button(
            dm_actions, text="📄 Сгенерировать PDF", command=self._generate_dm_pdf,
            style="Accent.TButton",
        )
        self.dm_generate_button.grid(row=0, column=0, padx=(0, 6), pady=2)
        self.dm_generate_tooltip = _make_tooltip(
            self.dm_generate_button, "Создать PDF с этикетками из выбранного CSV."
        )
        self.dm_openpdf_button = ttk.Button(
            dm_actions, text="📂 Открыть PDF", command=self._open_dm_pdf,
            style="Neutral.TButton"
        )
        self.dm_openpdf_button.grid(row=0, column=1, padx=(0, 6), pady=2)
        self.dm_preview_button = ttk.Button(
            dm_actions, text="👁 Предпросмотр", command=self._open_full_preview,
            style="Info.TButton"
        )
        self.dm_preview_button.grid(row=0, column=2, padx=(0, 6), pady=2)
        self.dm_print_button = ttk.Button(
            dm_actions, text="🖨 Печать", command=self._print_dm_pdf,
            style="Warn.TButton"
        )
        self.dm_print_button.grid(row=0, column=3, padx=(0, 6), pady=2)
        self.dm_openfolder_button = ttk.Button(
            dm_actions, text="📁 Открыть папку", command=self._open_dm_output_folder,
            style="Neutral.TButton"
        )
        self.dm_openfolder_button.grid(row=0, column=4, padx=(0, 6), pady=2)
        self.dm_checkcodes_button = ttk.Button(
            dm_actions, text="✔ Проверить коды", command=self._check_dm_codes,
            style="Info.TButton"
        )
        self.dm_checkcodes_button.grid(row=0, column=5, padx=(0, 6), pady=2)
        self.dm_exportpng_button = ttk.Button(
            dm_actions, text="🖼 Экспорт PNG", command=self._export_dm_png,
            style="Warn.TButton"
        )
        self.dm_exportpng_button.grid(row=0, column=6, padx=(0, 6), pady=2)

        self.dm_action_row = dm_actions

        # --- Строка: шаблоны этикеток (профили настроек) ---
        profile_row = ttk.Frame(output_frame)
        profile_row.grid(row=3, column=0, columnspan=5, sticky="ew", pady=(0, 8))
        profile_row.columnconfigure(1, weight=1)
        ttk.Label(profile_row, text="Шаблон:").grid(
            row=0, column=0, sticky="e", padx=(0, 8)
        )
        self.dm_profiles_combo = ttk.Combobox(
            profile_row, textvariable=self.dm_profile_name_var,
            state="readonly", width=18,
        )
        self.dm_profiles_combo.grid(row=0, column=1, sticky="ew", padx=(0, 8))
        self.dm_profiles_combo.bind("<<ComboboxSelected>>", self._apply_selected_profile)
        self.dm_save_profile_button = ttk.Button(
            profile_row, text="Сохранить шаблон", command=self._save_profile_dialog,
            style="Info.TButton"
        )
        self.dm_save_profile_button.grid(row=0, column=2, padx=(0, 6))
        self.dm_load_profile_button = ttk.Button(
            profile_row, text="Загрузить шаблон", command=self._load_profile_dialog,
            style="Neutral.TButton"
        )
        self.dm_load_profile_button.grid(row=0, column=3)
        self._refresh_profiles_combo()

        # внутренний журнал вкладки убран — всё идёт в общий «Журнал»

        ttk.Label(left, textvariable=self.dm_status_var, anchor="w").grid(
            row=2, column=0, sticky="ew", pady=(8, 0)
        )
        self.dm_progress = ttk.Progressbar(left, mode="determinate", maximum=100)
        self.dm_progress.grid(row=3, column=0, sticky="ew", pady=(4, 0))
        self.dm_progress_label = ttk.Label(left, text="", anchor="w")
        self.dm_progress_label.grid(row=4, column=0, sticky="ew")

        # ===== Правая часть: живой макет одной этикетки =====
        live_frame = ttk.LabelFrame(tab2_body, text="Макет этикетки (пример)")
        live_frame.grid(row=0, column=1, sticky="nsew", padx=(12, 0))
        live_frame.columnconfigure(0, weight=1)
        live_frame.rowconfigure(0, weight=1)

        self.dm_live_canvas = tk.Canvas(
            live_frame, bg="#ffffff", highlightthickness=0, height=340
        )
        self.dm_live_canvas.grid(row=0, column=0, sticky="nsew")
        self.dm_live_canvas.pack_propagate(False)

        # при первом построении интерфейса рисуем пустой макет
        self.root.after(100, self._update_live_mock)

        # --- Вкладка 3: Полный предпросмотр убран из вкладок — открывается
        # отдельным модальным окном по кнопке «Предпросмотр» (_open_full_preview).
        # Канвас для него создаётся внутри модального окна.

        # --- Вкладка 4: Быстрый DataMatrix (одиночный код для проверки) ---
        tab4 = ttk.Frame(self.notebook, padding=10)
        self.notebook.add(tab4, text="Быстрый DataMatrix")
        tab4.columnconfigure(0, weight=1)
        tab4.rowconfigure(3, weight=1)

        ttk.Label(
            tab4,
            text="Вставьте один код маркировки («Честный знак») — программа проверит "
                 "его структуру и покажет DataMatrix 25×25 мм в полном размере "
                 "(читается при наведении телефона).\n"
                 "Поддерживается сканер штрихкодов в клавиатурном режиме: введите "
                 "код в любое поле кода и нажмите Enter — программа перенесёт его "
                 "сюда и запустит проверку.",
        ).grid(row=0, column=0, columnspan=3, sticky="w", pady=(0, 8))

        qrow = ttk.Frame(tab4)
        qrow.grid(row=1, column=0, columnspan=3, sticky="ew", pady=(0, 8))
        qrow.columnconfigure(0, weight=1)
        ttk.Label(qrow, text="Код маркировки:").grid(
            row=0, column=0, sticky="w", pady=(0, 4))
        # многострочное поле — длинный код «Честного знака» не обрезается
        self.quick_code_text = scrolledtext.ScrolledText(
            qrow, height=4, wrap="none", width=70
        )
        self.quick_code_text.grid(row=1, column=0, sticky="ew", padx=(0, 6))
        self.quick_code_text.bind("<KeyRelease>", self._on_quick_code_change)
        self.quick_code_text.bind("<Control-Return>", self._quick_generate)
        ttk.Button(qrow, text="Проверить", command=self._quick_generate,
                   style="Accent.TButton").grid(
            row=2, column=0, sticky="w", padx=(0, 6), pady=(6, 0))
        # Enter запускает проверку/генерацию — так работают сканеры штрихкодов
        self.quick_code_text.bind("<Return>", self._quick_generate)

        ttk.Label(tab4, textvariable=self.quick_status_var).grid(
            row=2, column=0, columnspan=3, sticky="w", pady=(0, 8))
        self.quick_preview_canvas = tk.Canvas(
            tab4, bg="#ffffff", highlightthickness=0
        )
        self.quick_preview_canvas.grid(row=3, column=0, sticky="nsew")

        # --- Вкладка 5: Журнал (общий лог событий программы, как в 1С) ---
        tab5 = ttk.Frame(self.notebook, padding=10)
        self.notebook.add(tab5, text="Журнал")
        tab5.columnconfigure(0, weight=1)
        tab5.rowconfigure(1, weight=1)

        ttk.Label(
            tab5,
            text="Журнал событий программы: операции, статусы и ошибки со всех вкладок.",
        ).grid(row=0, column=0, sticky="w", pady=(0, 8))

        self.global_log = scrolledtext.ScrolledText(
            tab5, state="disabled", wrap="word", width=100, height=24
        )
        self.global_log.grid(row=1, column=0, sticky="nsew")
        # двойной клик по строке журнала — расшифровка найденного в ней кода
        self.global_log.bind("<Double-Button-1>", self._on_dm_log_double)

        log_btns = ttk.Frame(tab5)
        log_btns.grid(row=2, column=0, sticky="ew", pady=(8, 0))
        ttk.Button(
            log_btns, text="Очистить журнал",
            command=self._clear_global_log,
        ).pack(side="left")
        ttk.Label(
            log_btns, text="Двойной клик по строке — подробности кода (если это код маркировки)."
        ).pack(side="left", padx=12)

        # перетаскивание вкладок мышкой
        self._setup_tab_drag()
        # приём сканера штрихкодов (клавиатурный режим)
        self._setup_scanner()

        # после отрисовки всех элементов подгоняем размер окна под содержимое,
        # чтобы ни нижняя, ни правая граница не обрезали контент
        self.root.after(50, self._fit_window)

    def _fit_window(self) -> None:
        """Автоматически подстраивает размер окна, чтобы весь контент
        помещался без обрезания по краям."""
        try:
            self.root.update_idletasks()
            # требуемые размеры по содержимому
            req_w = self.root.winfo_reqwidth()
            req_h = self.root.winfo_reqheight()
            cur_w, cur_h = self.root.winfo_width(), self.root.winfo_height()
            # берём максимум: не меньше требуемого, но не больше в разумных пределах
            new_w = max(cur_w, min(req_w, 2000))
            new_h = max(cur_h, min(req_h, 1200))
            if new_w != cur_w or new_h != cur_h:
                self.root.geometry(f"{new_w}x{new_h}")
        except Exception:
            pass

    def _set_dm_csv_files(self, files: list[Path]) -> None:
        """Сохраняет список CSV-файлов и обновляет поле выбора и предпросмотр."""
        self._dm_csv_files = list(files)
        self._dm_cached_first_code = None
        if not files:
            return
        # кешируем первый код заранее, чтобы не перечитывать все файлы при каждом
        # обновлении живого макета (это и было причиной сильных тормозов)
        try:
            from converter import parse_csv_codes
            text = self._collect_dm_csv_text()
            codes = parse_csv_codes(text)
            self._dm_cached_first_code = codes[0] if codes else None
        except Exception:
            self._dm_cached_first_code = None
        if len(files) == 1:
            self.dm_csv_var.set(str(files[0]))
            self.dm_pdf_var.set(str(files[0].with_suffix(".pdf")))
        else:
            self.dm_csv_var.set(
                f"{len(files)} файлов: "
                + "; ".join(str(p) for p in files[:2])
                + (" ..." if len(files) > 2 else "")
            )
            self.dm_pdf_var.set(str(files[0].with_suffix(".pdf")))
        self._update_live_mock()

    def _schedule_preview(self, delay_ms: int = 30) -> None:
        """Отложенный запуск предпросмотра — чтобы канвас успел получить
        размер и не было дёргания рамки при перерисовке."""
        if hasattr(self, "_preview_after_id") and self._preview_after_id:
            try:
                self.root.after_cancel(self._preview_after_id)
            except Exception:
                pass
        self._preview_after_id = self.root.after(delay_ms, self._preview_dm)

    def _select_dm_csv(self) -> None:
        selected = filedialog.askopenfilenames(
            title="Выберите один или несколько CSV-файлов с кодами",
            filetypes=(("CSV files", "*.csv"), ("Text files", "*.txt"), ("All files", "*.*")),
        )
        if selected:
            self._set_dm_csv_files([Path(p) for p in selected])

    def _select_dm_folder(self) -> None:
        folder = filedialog.askdirectory(
            title="Выберите папку с CSV-файлами кодов"
        )
        if not folder:
            return
        root = Path(folder)
        files = [
            p for p in root.rglob("*")
            if p.is_file() and p.suffix.lower() in (".csv", ".txt")
        ]
        if not files:
            messagebox.showerror("Ошибка", f"В папке не найдено CSV/TXT-файлов:\n{folder}")
            return
        self._set_dm_csv_files(sorted(files, key=lambda p: str(p).casefold()))

    def _select_dm_pdf(self) -> None:
        selected = filedialog.asksaveasfilename(
            title="Куда сохранить PDF",
            initialdir=self.last_output_dir or str(Path.home()),
            initialfile="Коды_DataMatrix.pdf",
            defaultextension=".pdf",
            filetypes=(("PDF files", "*.pdf"), ("All files", "*.*")),
            confirmoverwrite=False,
        )
        if selected:
            self.dm_pdf_var.set(selected)

    def _open_dm_pdf(self) -> None:
        """Открывает сгенерированный PDF-файл."""
        if getattr(self, "_generating", False):
            messagebox.showinfo("Пожалуйста, подождите", "Идёт генерация PDF, дождитесь завершения.")
            return
        # открываем фактически созданный файл (мог появиться суффикс _1, _2)
        if getattr(self, "_last_generated_pdf", None) is not None:
            path = Path(self._last_generated_pdf)
            if path.is_file():
                try:
                    os.startfile(str(path))  # type: ignore[attr-defined]
                except Exception as exc:
                    messagebox.showerror("Ошибка", f"Не удалось открыть PDF:\n{exc}")
                return
        out_path = self.dm_pdf_var.get().strip()
        if not out_path:
            messagebox.showerror("Ошибка", "Сначала укажите путь к PDF-файлу")
            return
        path = Path(out_path)
        if not path.is_file():
            messagebox.showerror(
                "Ошибка",
                f"PDF-файл не найден:\n{path}\n\nСначала сгенерируйте PDF.",
            )
            return
        try:
            os.startfile(str(path))  # type: ignore[attr-defined]
        except Exception as exc:
            messagebox.showerror("Ошибка", f"Не удалось открыть PDF:\n{exc}")

    @staticmethod
    def _resolve_saved_pdf(output_pdf: Path) -> str:
        """Возвращает путь, по которому на самом деле сохранён PDF.

        Устаревший хелпер: реальный путь теперь возвращает сама функция
        generate_label_pdf (третьим значением), поэтому он больше не
        используется. Оставлен для обратной совместимости внешнего кода.
        """
        save_path = output_pdf
        n = 1
        while save_path.exists():
            save_path = output_pdf.with_stem(f"{output_pdf.stem}_{n}")
            n += 1
        return str(save_path)

    def _open_dm_output_folder(self) -> None:
        """Открывает папку, где сохраняется/был сохранён PDF."""
        out_path = self.dm_pdf_var.get().strip()
        if out_path:
            folder = Path(out_path).parent
            if folder.is_dir():
                os.startfile(str(folder))  # type: ignore[attr-defined]
                return
        # если путь не задан — открываем последнюю папку вывода или домашнюю
        folder = self.last_output_dir or str(Path.home())
        if os.path.isdir(folder):
            os.startfile(str(folder))  # type: ignore[attr-defined]
        else:
            messagebox.showerror("Ошибка", "Папка не найдена")

    def _export_dm_png(self) -> None:
        """Экспортирует каждую этикетку как отдельный PNG в выбранную папку."""
        from converter import export_labels_png, parse_csv_codes

        files = getattr(self, "_dm_csv_files", None) or []
        if not files:
            single = self.dm_csv_var.get().strip()
            if single:
                files = [Path(single)]
        if not files:
            messagebox.showerror("Ошибка", "Сначала выберите CSV-файл(ы)")
            return
        out_dir = filedialog.askdirectory(
            title="Куда сохранить PNG-этикетки",
            initialdir=self.last_output_dir or str(Path.home()),
        )
        if not out_dir:
            return

        bad_date = self._check_bad_date()
        if bad_date:
            messagebox.showerror("Ошибка", f"Неверная дата: {bad_date}")
            return

        try:
            size_w, size_h = self._parse_label_size(self.dm_label_size_var.get())
        except ValueError:
            messagebox.showerror("Ошибка", "Выберите корректный размер этикетки")
            return
        csv_text = self._collect_dm_csv_text()
        total_codes = len(parse_csv_codes(csv_text))

        self.dm_status_var.set("Экспорт PNG...")
        self.dm_exportpng_button.configure(state="disabled")
        self._generating = True
        self._cancel_event = threading.Event()
        self._dm_progress_state = [0, total_codes]
        self._open_progress_dialog("Экспорт PNG", total_codes)

        # Собираем настройки в главном потоке (Tk-переменные нельзя читать
        # из фонового потока)
        settings = {
            "size_w": size_w,
            "size_h": size_h,
            "heading": self.dm_name_var.get().strip(),
            "production_date": self.dm_date_var.get().strip(),
            "expiry": self.dm_expiry_var.get().strip(),
            "batch": self.dm_batch_var.get().strip(),
            "show_date": bool(self.dm_show_date_var.get()),
            "show_expiry": bool(self.dm_show_expiry_var.get()),
            "show_batch": bool(self.dm_show_batch_var.get()),
            "show_code": bool(self.dm_show_code_var.get()),
            "use_short": bool(self.dm_show_short_var.get()),
            "bold_name": bool(self.dm_bold_name_var.get()),
            "bold_date": bool(self.dm_bold_date_var.get()),
            "bold_expiry": bool(self.dm_bold_expiry_var.get()),
            "bold_batch": bool(self.dm_bold_batch_var.get()),
            "bold_code": bool(self.dm_bold_code_var.get()),
            "size_name": _float_or_none(self.dm_size_name_var.get()),
            "size_date": _float_or_none(self.dm_size_date_var.get()),
            "size_expiry": _float_or_none(self.dm_size_expiry_var.get()),
            "size_batch": _float_or_none(self.dm_size_batch_var.get()),
            "size_code": _float_or_none(self.dm_size_code_var.get()),
        }

        def work() -> None:
            try:
                saved, err = export_labels_png(
                    csv_text,
                    Path(out_dir),
                    label_w_mm=size_w,
                    label_h_mm=size_h,
                    heading=settings["heading"],
                    production_date=settings["production_date"],
                    expiry=settings["expiry"],
                    batch=settings["batch"],
                    show_date=settings["show_date"],
                    show_expiry=settings["show_expiry"],
                    show_batch=settings["show_batch"],
                    show_code=settings["show_code"],
                    use_short=settings["use_short"],
                    dpi=300,
                    bold_name=settings["bold_name"],
                    bold_date=settings["bold_date"],
                    bold_expiry=settings["bold_expiry"],
                    bold_batch=settings["bold_batch"],
                    bold_code=settings["bold_code"],
                    size_name=settings["size_name"],
                    size_date=settings["size_date"],
                    size_expiry=settings["size_expiry"],
                    size_batch=settings["size_batch"],
                    size_code=settings["size_code"],
                    progress_cb=lambda done, total: self._on_progress_state(done, total),
                    cancel_event=self._cancel_event,
                )
                self.events.put(("dm_png_exported", {
                    "saved": int(saved),
                    "err": err or "",
                    "total_codes": total_codes,
                    "out_dir": str(out_dir),
                }))
            except Exception as exc:
                self.events.put(("dm_png_exported", {
                    "saved": 0,
                    "err": f"Ошибка: {exc}" if str(exc) else "Ошибка экспорта PNG",
                    "total_codes": total_codes,
                    "out_dir": str(out_dir),
                }))

        threading.Thread(target=work, daemon=True).start()

    def _check_dm_codes(self) -> None:
        """Проверяет коды из выбранных CSV: структуру, дубликаты.
        Показывает результат в журнале и всплывающим окном."""
        from converter import parse_csv_codes
        try:
            csv_text = self._collect_dm_csv_text()
        except Exception as exc:
            messagebox.showerror("Ошибка", f"Не удалось прочитать CSV:\n{exc}")
            return
        codes = parse_csv_codes(csv_text)
        if not codes:
            messagebox.showinfo("Проверка кодов", "В выбранных CSV-файлах нет кодов.")
            return

        details = validate_codes(codes)
        total = len(details)
        ok = sum(1 for d in details if not d["problems"])
        bad = [d for d in details if d["problems"]]
        duplicates = sum(1 for d in details if d["is_duplicate"])

        lines = ["=== Проверка кодов ===", f"Всего кодов: {total}", f"Корректных: {ok}",
                 f"С проблемами: {len(bad)}", f"Дубликатов: {duplicates}"]
        if bad:
            lines.append("")
            lines.append("Проблемные коды:")
            for d in bad:
                detail = "; ".join(d["problems"])
                lines.append(f"  №{d['index']}: {detail}")
                self._log_op(f"Проверка: код №{d['index']} — {detail}")
        text = "\n".join(lines)

        self._global_log(text)
        self.dm_status_var.set(f"Проверка: {ok} из {total} корректных")

        if bad:
            messagebox.showwarning(
                "Проверка кодов",
                f"Всего кодов: {total}\nКорректных: {ok}\n"
                f"С проблемами: {len(bad)}\nДубликатов: {duplicates}\n\n"
                f"{text.split(chr(10) + chr(10), 1)[1]}\n\nПодробности в журнале.",
            )
        else:
            messagebox.showinfo(
                "Проверка кодов",
                f"Всего кодов: {total}\nКорректных: {ok}\nВсе коды валидны.",
            )
        self._log_op(
            f"Проверка кодов: всего {total}, корректных {ok}, "
            f"с проблемами {len(bad)}, дублей {duplicates}"
        )

    def _decode_code_fields(self, code: str) -> list[tuple[str, str]]:
        """Разбирает полный код маркировки на поля AI 01/21/91/92.
        Возвращает список (название, значение)."""
        GS_CHAR = "\x1d"
        fields: list[tuple[str, str]] = []
        rest = code
        if rest.startswith("01"):
            gtin = rest[2:16]
            rest = rest[16:]
            fields.append(("GTIN (01)", gtin))
        if rest.startswith("21"):
            serial = rest[2:]
            if GS_CHAR in serial:
                serial = serial.split(GS_CHAR, 1)[0]
            fields.append(("Серийный номер (21)", serial))
        for part in code.split(GS_CHAR):
            if part.startswith("91"):
                fields.append(("Криптоключ (91)", part[2:]))
            elif part.startswith("92"):
                fields.append(("Цифровая подпись (92)", part[2:]))
        return fields

    def _show_code_details(self, code: str) -> None:
        """Открывает окно с расшифровкой полей кода маркировки."""
        code = (code or "").strip().strip('"')
        if not code:
            messagebox.showinfo("Расшифровка кода", "Введите/выберите код маркировки.")
            return
        fields = self._decode_code_fields(code)
        if not fields:
            messagebox.showinfo(
                "Расшифровка кода",
                "Не удалось распознать структуру кода.\nОжидается 01+GTIN+21+серия(+91+92).",
            )
            return

        dialog = tk.Toplevel(self.root)
        dialog.title("Расшифровка кода")
        dialog.geometry("720x420")
        dialog.minsize(480, 260)
        dialog.transient(self.root)
        dialog.grab_set()

        dark = self.dark_var.get()
        bg = "#1f1f1f" if dark else "#f0f0f0"
        fg = "#e8e8e8" if dark else "#000000"
        dialog.configure(bg=bg)

        body = ttk.Frame(dialog)
        body.pack(fill="both", expand=True, padx=12, pady=10)
        body.columnconfigure(1, weight=1)

        ttk.Label(body, text="Полный код:").grid(row=0, column=0, sticky="nw",
                                                 padx=(0, 8), pady=(0, 8))
        self._detail_code_var = tk.StringVar(value=code)
        code_entry = ttk.Entry(body, textvariable=self._detail_code_var)
        code_entry.grid(row=0, column=1, sticky="ew", pady=(0, 8))

        for i, (label, value) in enumerate(fields, start=1):
            ttk.Label(body, text=label + ":").grid(
                row=i, column=0, sticky="nw", padx=(0, 8), pady=2
            )
            val_entry = ttk.Entry(body)
            val_entry.insert(0, value)
            val_entry.configure(state="disabled")
            val_entry.grid(row=i, column=1, sticky="ew", pady=2)

        buttons = ttk.Frame(dialog)
        buttons.pack(fill="x", padx=12, pady=(0, 10))

        def copy_code() -> None:
            self.root.clipboard_clear()
            self.root.clipboard_append(self._detail_code_var.get())
            self.dm_status_var.set("Код скопирован")

        ttk.Button(buttons, text="Копировать код", command=copy_code).pack(side="left")
        ttk.Button(buttons, text="Закрыть", command=dialog.destroy).pack(
            side="left", padx=(8, 0)
        )
        dialog.after(10, lambda: dialog.lift())

    def _show_details_for_line(self, line: str) -> None:
        """По клику/двойному клику по строке журнала открывает расшифровку,
        если в строке найден код маркировки."""
        for token in line.split():
            if token.startswith("01") and len(token) >= 16:
                self._show_code_details(token)
                return

    def _on_dm_log_double(self, event) -> None:
        """Обрабатывает двойной клик по журналу: открывает расшифровку кода
        из строки под курсором, если там есть код маркировки."""
        try:
            index = self.global_log.index(f"@{event.x},{event.y}")
            line = self.global_log.get(f"{index} linestart", f"{index} lineend")
        except Exception:
            line = ""
        if line:
            self._show_details_for_line(line)

    def _manual_code_generate(self) -> None:
        """Генерирует DataMatrix по вручную введённому коду маркировки.

        Переносит код из поля «Код маркировки» во вкладку «Быстрый DataMatrix»
        и запускает генерацию там (механизм тот же, что у кнопки «Проверить»
        в «Быстром DataMatrix»).
        """
        code = getattr(self, "dm_manual_code_var", None)
        if code is None:
            return
        raw = code.get().strip()
        if not raw:
            messagebox.showwarning(
                "Ручной ввод", "Введите код маркировки в поле «Код маркировки».")
            return
        self._set_quick_code(raw)
        self._select_tab("Быстрый DataMatrix")
        self._quick_generate()

    def _collect_dm_csv_text(self) -> str:
        """Возвращает объединённый текст кодов из всех выбранных CSV-файлов."""
        files = getattr(self, "_dm_csv_files", None) or []
        if not files:
            single = self.dm_csv_var.get().strip()
            if single:
                files = [Path(single)]
        parts = []
        for f in files:
            text = f.read_text(encoding="utf-8-sig", errors="replace").rstrip("\r\n")
            if text:
                parts.append(text)
        return "\n".join(parts)

    def _check_bad_date(self) -> str:
        """Возвращает текст ошибки, если поле даты производства или срока
        годности заполнено и невалидно; подсвечивает поля красным.
        Иначе возвращает пустую строку."""
        self._mark_date_status()
        self._mark_expiry_status()
        reason = self._invalid_date_reason() or self._invalid_expiry_reason()
        return reason or ""

    def _generate_dm_pdf(self) -> None:
        out_path = self.dm_pdf_var.get().strip()
        files = getattr(self, "_dm_csv_files", None) or []
        if not files:
            single = self.dm_csv_var.get().strip()
            if single:
                files = [Path(single)]
        if not files:
            messagebox.showerror("Ошибка", "Не выбран ни один CSV-файл")
            return
        if not out_path:
            messagebox.showerror("Ошибка", "Не указан путь для PDF")
            return
        out_path = Path(out_path)
        missing = [str(f) for f in files if not f.is_file()]
        if missing:
            messagebox.showerror(
                "Ошибка", f"CSV-файл не найден:\n" + "\n".join(missing)
            )
            return
        bad_date = self._check_bad_date()
        if bad_date:
            messagebox.showerror("Ошибка", f"Неверная дата: {bad_date}")
            return

        try:
            size_w, size_h = self._parse_label_size(self.dm_label_size_var.get())
        except ValueError:
            messagebox.showerror("Ошибка", "Выберите корректный размер этикетки")
            return

        def _append_log(text: str) -> None:
            self._global_log(str(text).rstrip())

        try:
            csv_text = self._collect_dm_csv_text()
        except Exception as exc:
            messagebox.showerror("Ошибка", f"Не удалось прочитать CSV:\n{exc}")
            return

        from converter import parse_csv_codes
        all_codes = parse_csv_codes(csv_text)
        total_codes = len(all_codes)

        self.dm_status_var.set(f"Генерация (всего: {total_codes})...")
        self.dm_generate_button.configure(state="disabled")
        self._generating = True
        self._cancel_event = threading.Event()
        self._dm_progress_state = [0, total_codes]  # (done, total) обновляется из потока
        self.dm_progress.configure(value=0)
        self.dm_progress_label.configure(text=f"0 из {total_codes}")
        self._open_progress_dialog("Генерация PDF", total_codes)

        def _poll_progress() -> None:
            if not self._generating:
                return
            done, total = self._dm_progress_state
            pct = int(done * 100.0 / total) if total else 0
            self.dm_progress.configure(value=pct)
            self.dm_progress_label.configure(text=f"Обработано: {done} из {total}")
            self.root.after(100, _poll_progress)
        self.root.after(100, _poll_progress)

        # Собираем в ГЛАВНОМ потоке все значения настроек и диагностику:
        # обращаться к Tk-переменным или вызывать root.after из фонового
        # потока нельзя (RuntimeError: main thread is not in main loop).
        settings = {
            "size_w": size_w,
            "size_h": size_h,
            "scale": self._dm_scale_for(size_w),
            "heading": self.dm_name_var.get().strip(),
            "production_date": self.dm_date_var.get().strip(),
            "expiry": self.dm_expiry_var.get().strip(),
            "batch": self.dm_batch_var.get().strip(),
            "show_date": bool(self.dm_show_date_var.get()),
            "show_expiry": bool(self.dm_show_expiry_var.get()),
            "show_batch": bool(self.dm_show_batch_var.get()),
            "show_code": bool(self.dm_show_code_var.get()),
            "use_short": bool(self.dm_show_short_var.get()),
            "bold_name": bool(self.dm_bold_name_var.get()),
            "bold_date": bool(self.dm_bold_date_var.get()),
            "bold_expiry": bool(self.dm_bold_expiry_var.get()),
            "bold_batch": bool(self.dm_bold_batch_var.get()),
            "bold_code": bool(self.dm_bold_code_var.get()),
            "size_name": _float_or_none(self.dm_size_name_var.get()),
            "size_date": _float_or_none(self.dm_size_date_var.get()),
            "size_expiry": _float_or_none(self.dm_size_expiry_var.get()),
            "size_batch": _float_or_none(self.dm_size_batch_var.get()),
            "size_code": _float_or_none(self.dm_size_code_var.get()),
            "diag": {
                "size_mm": [size_w, size_h],
                "production_date": self.dm_date_var.get().strip(),
                "expiry": self.dm_expiry_var.get().strip(),
                "batch": self.dm_batch_var.get().strip(),
                "show_date": bool(self.dm_show_date_var.get()),
                "show_expiry": bool(self.dm_show_expiry_var.get()),
                "show_batch": bool(self.dm_show_batch_var.get()),
                "show_code": bool(self.dm_show_code_var.get()),
                "use_short": bool(self.dm_show_short_var.get()),
                "codes": total_codes,
            },
        }

        def work() -> None:
            try:
                csv_text = self._collect_dm_csv_text()
                def _on_progress(done, total):
                    self._on_progress_state(done, total)
                pages, err, save_path = generate_label_pdf(
                    csv_text,
                    out_path,
                    label_w_mm=size_w,
                    label_h_mm=size_h,
                    heading=settings["heading"],
                    production_date=settings["production_date"],
                    expiry=settings["expiry"],
                    batch=settings["batch"],
                    show_date=settings["show_date"],
                    show_expiry=settings["show_expiry"],
                    show_batch=settings["show_batch"],
                    show_code=settings["show_code"],
                    use_short=settings["use_short"],
                    scale=settings["scale"],
                    bold_name=settings["bold_name"],
                    bold_date=settings["bold_date"],
                    bold_expiry=settings["bold_expiry"],
                    bold_batch=settings["bold_batch"],
                    bold_code=settings["bold_code"],
                    size_name=settings["size_name"],
                    size_date=settings["size_date"],
                    size_expiry=settings["size_expiry"],
                    size_batch=settings["size_batch"],
                    size_code=settings["size_code"],
                    progress_cb=_on_progress,
                    cancel_event=self._cancel_event,
                )
                # реальный путь сохранения (учитывает копии с номером вместо
                # перезаписи, если файл уже существовал) — напрямую ОТ generator
                actual = save_path or out_path
                # Результат передаём в главный поток через очередь событий;
                # там же обновляем GUI, статус и авто-открытие.
                self.events.put(("dm_generated", {
                    "pages": pages,
                    "err": err or "",
                    "save_path": str(actual),
                    "total_codes": total_codes,
                    "out_path": str(out_path),
                    "diag": settings["diag"],
                }))
            except Exception as exc:
                self.events.put(("dm_generated", {
                    "pages": 0,
                    "err": f"Ошибка: {exc}" if str(exc) else "Ошибка генерации PDF",
                    "save_path": str(out_path),
                    "total_codes": total_codes,
                    "out_path": str(out_path),
                    "diag": settings["diag"],
                }))

        threading.Thread(target=work, daemon=True).start()

    def _on_progress_state(self, done: int, total: int) -> None:
        """Обновляет состояние прогресса (безопасно вызывается из потока)."""
        self._dm_progress_state[0] = done
        self._dm_progress_state[1] = total

    def _open_progress_dialog(self, title: str, total: int) -> None:
        """Показывает модальное окно прогресса генерации с кнопкой «Отмена».

        Создаётся дочернее модальное окно (grab_set) с прогресс-баром и кнопкой
        «Отмена», которая ставит self._cancel_event. Окно закрывается вызовом
        self._close_progress_dialog() по завершении генерации.
        """
        if hasattr(self, "_progress_win") and self._progress_win is not None:
            try:
                self._progress_win.destroy()
            except Exception:
                pass
        win = tk.Toplevel(self.root)
        win.title(title)
        win.resizable(False, False)
        win.transient(self.root)
        win.grab_set()
        win.protocol("WM_DELETE_WINDOW", self._cancel_and_close_progress)
        ttk.Label(win, text=title).pack(padx=16, pady=(12, 4))
        bar = ttk.Progressbar(win, mode="determinate", maximum=1000, length=360)
        bar.pack(padx=16, pady=4)
        info = ttk.Label(win, text="Подготовка…")
        info.pack(padx=16, pady=4)
        btns = ttk.Frame(win)
        btns.pack(pady=8)
        cancel_btn = ttk.Button(btns, text="Отмена", style="Danger.TButton",
                                command=self._cancel_and_close_progress)
        cancel_btn.pack()
        self._progress_win = win
        self._progress_bar = bar
        self._progress_info = info
        self._progress_total = total or 1
        self._progress_done = 0
        self._poll_progress_dialog()

    def _poll_progress_dialog(self) -> None:
        """Обновляет прогресс-бар модального окна (вызывается по таймеру)."""
        win = getattr(self, "_progress_win", None)
        if win is None or not win.winfo_exists():
            return
        done, total = self._dm_progress_state
        self._progress_done = done
        self._progress_total = total or 1
        pct = int(done * 1000.0 / self._progress_total)
        try:
            self._progress_bar.configure(value=pct)
            self._progress_info.configure(text=f"Обработано: {done} из {total}")
        except Exception:
            pass
        self.root.after(100, self._poll_progress_dialog)

    def _cancel_and_close_progress(self) -> None:
        """Кнопка «Отмена»: ставит флаг отмены и закрывает модальное окно."""
        try:
            self._cancel_event.set()
        except Exception:
            pass
        self._close_progress_dialog()

    def _close_progress_dialog(self) -> None:
        """Закрывает окно прогресса и снимает блокировку главного окна."""
        win = getattr(self, "_progress_win", None)
        if win is None:
            return
        try:
            if win.winfo_exists():
                win.grab_release()
                win.destroy()
        except Exception:
            pass
        self._progress_win = None

    def _parse_label_size(self, label: str) -> tuple[float, float]:
        """Разбирает строку вида '58 × 40' в (ширина, высота) в мм."""
        label = label.replace("×", "x").replace("*", "x").replace(" ", "")
        parts = label.split("x")
        if len(parts) != 2:
            raise ValueError(f"Неверный размер: {label!r}")
        w = float(parts[0])
        h = float(parts[1])
        if w <= 0 or h <= 0:
            raise ValueError("Размер этикетки должен быть положительным")
        return w, h

    def _dm_scale_for(self, size_w: float) -> int:
        """Масштаб Zint (множитель модуля) под ширину этикетки, чтобы
        DataMatrix был достаточно крупным (не меньше 10×10 мм) для сканирования."""
        # размер модуля 0.254 мм => при ширине этикетки подкручиваем масштаб
        # ориентировочно: на 40 мм ширины ~ scale 4, на 58 мм ~ scale 6
        if size_w <= 40:
            return 4
        if size_w <= 50:
            return 5
        if size_w <= 60:
            return 6
        return 8

    def _on_dm_group_change(self, _event=None) -> None:
        """Подсказка по требуемому размеру/элементу кода и подбор размеров."""
        group = self.dm_group_var.get()
        hint = {
            "Медизделия": "Мин. элемент кода 0,255–0,680 мм",
            "Духи/парфюмерия": "Мин. элемент кода 0,255–0,615 мм",
            "Косметика": "Мин. элемент кода 0,255–0,615 мм; удобен формат 25×25",
            "Лёгкая промышленность": "Мин. элемент кода 0,255–0,680 мм (ППРФ №1956)",
            "Антисептики": "Мин. элемент кода 0,255–0,615 мм",
        }.get(group)
        if hasattr(self, "dm_size_combo"):
            sizes = self._sizes_for_group(group)
            self.dm_size_combo.configure(values=sizes)
            if self.dm_label_size_var.get() not in sizes:
                self.dm_label_size_var.set(sizes[0] if sizes else "58 × 40")
        if hint:
            self.dm_status_var.set(hint)
        self._update_live_mock()

    def _print_dm_pdf(self) -> None:
        """Печатает этикетки: строит PDF во временный файл и отправляет на
        печать штатным диалогом Windows."""
        from converter import generate_label_pdf, print_pdf_windows

        files = getattr(self, "_dm_csv_files", None) or []
        if not files:
            single = self.dm_csv_var.get().strip()
            if single:
                files = [Path(single)]
        if not files:
            messagebox.showerror("Ошибка", "Сначала выберите CSV-файл(ы)")
            return
        missing = [str(f) for f in files if not f.is_file()]
        if missing:
            messagebox.showerror("Ошибка", f"CSV-файл не найден:\n" + "\n".join(missing))
            return
        bad_date = self._check_bad_date()
        if bad_date:
            messagebox.showerror("Ошибка", f"Неверная дата: {bad_date}")
            return

        try:
            size_w, size_h = self._parse_label_size(self.dm_label_size_var.get())
        except ValueError:
            messagebox.showerror("Ошибка", "Выберите корректный размер этикетки")
            return

        self.dm_status_var.set("Подготовка к печати...")
        self.dm_print_button.configure(state="disabled")

        # Собираем настройки в главном потоке (Tk-переменные нельзя читать
        # из фонового потока)
        settings = {
            "size_w": size_w,
            "size_h": size_h,
            "scale": self._dm_scale_for(size_w),
            "heading": self.dm_name_var.get().strip(),
            "production_date": self.dm_date_var.get().strip(),
            "expiry": self.dm_expiry_var.get().strip(),
            "batch": self.dm_batch_var.get().strip(),
            "show_date": bool(self.dm_show_date_var.get()),
            "show_expiry": bool(self.dm_show_expiry_var.get()),
            "show_batch": bool(self.dm_show_batch_var.get()),
            "show_code": bool(self.dm_show_code_var.get()),
            "use_short": bool(self.dm_show_short_var.get()),
        }

        def work() -> None:
            import tempfile
            import traceback
            tmp_pdf = Path(tempfile.gettempdir()) / "delfin_dm_print.pdf"
            try:
                # удаляем временный PDF от прошлой печати, чтобы get save_path
                # не ушёл в _1, _2 и т.д.
                if tmp_pdf.exists():
                    tmp_pdf.unlink()
                csv_text = self._collect_dm_csv_text()
                pages, err, _save_path = generate_label_pdf(
                    csv_text,
                    tmp_pdf,
                    label_w_mm=size_w,
                    label_h_mm=size_h,
                    heading=settings["heading"],
                    production_date=settings["production_date"],
                    expiry=settings["expiry"],
                    batch=settings["batch"],
                    show_date=settings["show_date"],
                    show_expiry=settings["show_expiry"],
                    show_batch=settings["show_batch"],
                    show_code=settings["show_code"],
                    use_short=settings["use_short"],
                    scale=settings["scale"],
                )
                self.events.put(("dm_printed", {
                    "pages": pages,
                    "err": err or "",
                    "tmp_pdf": str(tmp_pdf),
                }))
            except Exception as exc:
                self.events.put(("dm_printed", {
                    "pages": 0,
                    "err": f"Ошибка: {exc}" if str(exc) else "Ошибка печати",
                    "tmp_pdf": str(tmp_pdf),
                }))

        threading.Thread(target=work, daemon=True).start()

    def _clean_code_input(self, raw: str) -> str:
        """Нормализует введённый код: убирает пробелы вокруг и символ разделителя,
        который пользователь мог вставить как переведённый GS (Unicode) или ASCII."""
        code = raw.strip()
        # замена визуальных вариантов разделителя на стандартный GS (\x1d)
        code = code.replace("\u29f8", "\x1d").replace("<>", "\x1d").replace("<GS>", "\x1d")
        return code

    def _quick_code_get(self) -> str:
        """Читает введённый код из многострочного поля (берёт первую строку)."""
        if not hasattr(self, "quick_code_text"):
            return self.quick_code_var.get()
        text = self.quick_code_text.get("1.0", "end-1c")
        return text.splitlines()[0] if text.splitlines() else ""

    def _on_quick_code_change(self, _event: object = None) -> None:
        """При вводе кода сбрасываем статус и убираем прежний предпросмотр."""
        self.quick_status_var.set("Вставьте код маркировки и нажмите «Проверить»")
        self.quick_preview_canvas.delete("all")
        self._quick_image_ref = None

    def _quick_clear(self) -> None:
        """Очищает поле ввода, статус и предпросмотр."""
        if hasattr(self, "quick_code_text"):
            self.quick_code_text.delete("1.0", "end")
        else:
            self.quick_code_var.set("")
        self.quick_status_var.set("Вставьте код маркировки и нажмите «Проверить»")
        self.quick_preview_canvas.delete("all")
        self._quick_image_ref = None

    def _quick_generate(self) -> None:
        """Проверяет введённый код и показывает DataMatrix 25×25 мм в 100% масштабе."""
        from converter import validate_code_structure, gs1data_to_zint_from_short, find_zint_exe
        code = self._clean_code_input(self._quick_code_get())
        if not code:
            self.quick_status_var.set("Пустой код. Вставьте код маркировки.")
            return
        problems = validate_code_structure(code)
        if problems:
            self.quick_status_var.set("Код не прошёл проверку: " + "; ".join(problems))
        else:
            self.quick_status_var.set("Код корректен. Генерирую DataMatrix…")
        self.root.update_idletasks()

        # строим DataMatrix из укороченного кода (01+GTIN+21+серия+91), потому что
        # именно так код считывается при проверке в «Честном знаке».
        try:
            zint_exe = find_zint_exe()
            if zint_exe is None:
                raise RuntimeError("Не найден zint.exe")
            zint_data = gs1data_to_zint_from_short(code)
            import subprocess
            import tempfile
            with tempfile.TemporaryDirectory() as td:
                png_path = Path(td) / "quick.png"
                # --scale=10 делает DataMatrix 25 мм при 100%, без сжатия вниз
                proc = subprocess.run(
                    [str(zint_exe), "-b", "DATAMATRIX", "--gs1", "--gssep",
                     "--square", "--scale=10", "-d", zint_data, "-o", str(png_path)],
                    capture_output=True, text=True,
                    creationflags=getattr(subprocess, "CREATE_NO_WINDOW", 0),
                )
                if proc.returncode != 0 or not png_path.exists():
                    raise RuntimeError("Ошибка генерации DataMatrix")
                from PIL import Image, ImageTk
                img = Image.open(str(png_path)).convert("RGBA")
                self._show_quick_image(img)
            self.quick_status_var.set(
                "Готово: DataMatrix 25×25 мм в полном размере. Наведите камеру телефона."
            )
        except Exception as exc:
            self.quick_status_var.set("Ошибка: " + str(exc))

    def _show_quick_image(self, img: object) -> None:
        """Отображает готовое изображение DataMatrix на канвасе в полном размере
        (без уменьшения), чтобы код было легко навести камерой."""
        from PIL import ImageTk
        photo = ImageTk.PhotoImage(img)
        self._quick_image_ref = photo  # держим ссылку, иначе фото исчезнет
        self.quick_preview_canvas.delete("all")
        self.quick_preview_canvas.create_image(0, 0, image=photo, anchor="nw")
        self.quick_preview_canvas.configure(scrollregion=(0, 0, img.width, img.height))

    def _toggle_tab(self, tab_text: str) -> None:
        """Показывает/скрывает вкладку по галочке в меню «Вид».

        Нельзя скрыть последнюю видимую вкладку — иначе окно останется пустым.
        """
        var = self._tab_visible.get(tab_text)
        if var is None:
            return
        try:
            shown = [t for t in self._tab_order if self._tab_visible[t].get()]
            if not var.get() and len(shown) <= 1:
                # не даём снять последнюю видимую вкладку
                var.set(True)
                return
            # находим id вкладки по тексту
            tab_id = None
            for tid in self.notebook.tabs():
                if self.notebook.tab(tid, "text") == tab_text:
                    tab_id = tid
                    break
            if tab_id is None:
                return
            if var.get():
                # показываем: вставляем обратно в правильное место по порядку
                self.notebook.add(tab_id)
                self.notebook.select(tab_id)
            else:
                self.notebook.hide(tab_id)
        except Exception:
            pass

    def _setup_scanner(self) -> None:
        """Готовит приём сканера штрихкодов (клавиатурный режим).

        Сканер работает как обычная клавиатура: вводит код в активное поле и
        шлёт Enter. Чтобы сканер работал «в любом поле кода», слушаем Enter
        на корневом окне: если это не поле «Быстрый DataMatrix», то введённый
        в активное поле текст переносим туда и запускаем генерацию. Это делает
        работу сканера удобной из любой точки окна.
        """
        if not hasattr(self, "root"):
            return

        self._scanner_on_enter = self._make_scanner_enter_handler()
        self._apply_scanner_enabled()

    def _make_scanner_enter_handler(self):
        """Создаёт и возвращает обработчик Enter сканера (для вкл/выкл приёма)."""
        def _on_enter(event=None) -> None:
            # если активно поле быстрого DataMatrix — там Enter уже обрабатывается
            try:
                focus = self.root.focus_get()
                if focus is getattr(self, "quick_code_text", None):
                    self._quick_generate()
                    return "break"
            except Exception:
                pass
            # если активное поле имеет какой-то текст — переносим его во вкладку
            try:
                if focus is not None and hasattr(focus, "get"):
                    try:
                        val = str(focus.get()).strip()
                    except Exception:
                        val = ""
                elif focus is not None and hasattr(focus, "get_view"):
                    try:
                        val = str(focus.get("1.0", "end-1c")).strip()
                    except Exception:
                        val = ""
                else:
                    val = ""
                if val and self._looks_like_code(val):
                    self._on_scanner_code(val)
                    return "break"
            except Exception:
                pass
            return None
        self._scanner_counter: int = 0
        return _on_enter

    def _apply_scanner_enabled(self) -> None:
        """Включает/выключает приём сканера по галочке в меню.

        В зависимости от выбранного режима:
          * keyboard — глобальный перехват Enter (сканер как клавиатура);
          * com — фоновый наблюдатель на COM-порту (сканер как Serial).
        """
        if not hasattr(self, "root"):
            return
        if not self.scanner_enabled.get():
            self._stop_com_listener()
            try:
                self.root.unbind("<Return>", self._scanner_on_enter)
            except Exception:
                pass
            return
        mode = getattr(self, "scanner_mode_var", None)
        if mode is not None and mode.get() == "com":
            self.root.unbind("<Return>", self._scanner_on_enter)
            self._start_com_listener()
        else:
            self._stop_com_listener()
            self.root.bind("<Return>", self._scanner_on_enter, add="+")

    def _start_com_listener(self) -> None:
        """Запускает фоновый наблюдатель на выбранном COM-порту."""
        import scanner as scanner_mod
        self._stop_com_listener()
        port = getattr(self, "scanner_port_var", None)
        if port is None or not port.get():
            return
        try:
            sc = scanner_mod.SerialScanner(port.get(), on_code=self._on_scanner_code,
                                           on_error=lambda m: self._scanner_com_error(m))
            err = sc.open()
            if err:
                self._scanner_com_error(err)
                return
            self._scanner_com = sc
        except Exception:
            self._scanner_com = None

    def _stop_com_listener(self) -> None:
        sc = getattr(self, "_scanner_com", None)
        if sc is not None:
            try:
                sc.close()
            except Exception:
                pass
        self._scanner_com = None

    def _scanner_com_error(self, message: str) -> None:
        self.root.after(0, lambda: self.status_var.set(message))

    def _on_scanner_code(self, val: str) -> None:
        """Сканированный код: инкремент счётчика и перенос в Быстрый DataMatrix."""
        self._scanner_counter = getattr(self, "_scanner_counter", 0) + 1
        self._set_quick_code(val)
        self._select_tab("Быстрый DataMatrix")
        self._quick_generate()

    def _show_scanner_check(self) -> None:
        """Окно «Сканер: настройки и проверка».

        Показывает, какое оборудование Windows видит как сканер (COM-порты
        и USB/HID-устройства ввода), позволяет выбрать режим приёма
        (Клавиатура или конкретный COM-порт) и проверить его, а также
        содержит поле со счётчиком считанных кодов.
        """
        import scanner_devices
        import scanner as scanner_mod

        win = tk.Toplevel(self.root)
        win.title("Сканер: настройки и проверка")
        win.resizable(False, False)
        win.transient(self.root)

        result_var = tk.StringVar(value="Наведите сканер на код…")
        counter_var = tk.StringVar(value="Считано кодов: 0")
        done = {"n": 0}
        status_var = tk.StringVar(value="")
        device_var = tk.StringVar(value="")
        com_evt = {"scanner": None}

        # ---------- обнаружение оборудования ----------
        hid_items = scanner_devices.list_hid_keyboards()
        com_items = scanner_devices.list_com_ports()

        tk.Label(win, text="Найденные сканеры:",
                 font=("Arial", 10, "bold")).pack(anchor="w", padx=16, pady=(12, 2))
        tk.Label(win, text="Выберите устройство, которым пользуетесь. "
                           "Мыши, клавиатуры и другая периферия не показываются.",
                 wraplength=440, justify="left", fg="#666"
                 ).pack(anchor="w", padx=16)

        # модель выбора устройства: (метка, ключ)
        device_list = ttk.Treeview(win, columns=("type",), show="tree",
                                   height=5, selectmode="browse")
        device_list.heading("#0", text="Устройство")
        device_list.column("#0", width=420)
        device_list.column("type", width=10, stretch=False)
        device_list.pack(padx=16, pady=(6, 2), fill="x")

        def _collect_devices() -> list[tuple[str, str, str]]:
            """Возвращает [(метка, тип)], где тип 'keyboard' или 'com:ПОРТ'."""
            items = []
            for p in com_items:
                label = f"{p.name or p.device}  [{p.device}]"
                items.append((label, f"com:{p.device}"))
            for h in hid_items:
                mark = "  (сканер)" if scanner_devices.is_likely_scanner_vid(h.vid) else ""
                label = f"{h.name}  (VID:{h.vid} PID:{h.pid}){mark}"
                items.append((label, "keyboard"))
            return items

        def _refresh_devs() -> None:
            nonlocal hid_items, com_items
            hid_items = scanner_devices.list_hid_keyboards()
            com_items = scanner_devices.list_com_ports()
            device_list.delete(*device_list.get_children())
            for label, key in _collect_devices():
                device_list.insert("", "end", iid=key, text=label,
                                   values=("usb",))
            # если выбранное устройство пропало (отключено) — сбрасываем
            if device_var.get():
                has = device_var.get() in {k for _l, k in _collect_devices()}
                if not has:
                    device_var.set("")
            if not device_list.get_children():
                status_var.set(
                    "Сканер не обнаружен. Подключите его по USB и нажмите «Обновить».")
            else:
                status_var.set("Найдено устройств: "
                               + str(len(device_list.get_children())))
                # предвыбор первого
                if not device_var.get():
                    first = device_list.get_children()[0]
                    device_list.selection_set(first)
                    device_var.set(first)

        def _on_select(_event=None) -> None:
            sel = device_list.selection()
            if sel:
                device_var.set(sel[0])

        device_list.bind("<<TreeviewSelect>>", _on_select)

        _refresh_devs()

        ttk.Label(win, textvariable=status_var, foreground="#a33",
                  wraplength=440, justify="left"
                  ).pack(anchor="w", padx=16, pady=(4, 0))

        # ---------- тест ----------
        sep = ttk.Separator(win, orient="horizontal")
        sep.pack(fill="x", padx=16, pady=(10, 6))
        ttk.Label(win, text="Проверка приёма:",
                  font=("Arial", 10, "bold")).pack(anchor="w", padx=16,
                                                   pady=(0, 4))

        ttk.Label(win, text="Клавиатурный режим: наведите сканер — код "
                            "появится ниже. COM-режим: откройте наблюдатель "
                            "и наведите сканер.",
                  wraplength=430, justify="left").pack(anchor="w", padx=16)

        entry = ttk.Entry(win, width=50)
        entry.pack(padx=16, pady=(6, 0))

        def _apply_code(code: str) -> bool:
            if not code or not self._looks_like_code(code):
                return False
            done["n"] += 1
            counter_var.set(f"Считано кодов: {done['n']}")
            result_var.set(f"Принято: {self._short_code(code)}")
            entry.delete(0, "end")
            return True

        def _on_local_enter(event=None) -> None:
            widget = self.root.focus_get()
            if widget is not entry:
                return None
            code = entry.get()
            if _apply_code(code):
                return "break"
            return None

        entry.bind("<Return>", _on_local_enter)

        # возвращает выбранное устройство: ("keyboard", "") или ("com", "COM3")
        def _selected_device() -> tuple[str, str]:
            key = device_var.get()
            if key.startswith("com:"):
                return ("com", key.split(":", 1)[1])
            return ("keyboard", "")

        # наблюдатель COM для теста (только если выбрано COM-устройство)
        def _start_watch() -> None:
            sc = com_evt.get("scanner")
            if sc is not None:
                try:
                    sc.close()
                except Exception:
                    pass
                com_evt["scanner"] = None
            mode, port = _selected_device()
            if mode != "com" or not port:
                status_var.set("Выберите COM-устройство из списка сканеров, "
                               "чтобы проверить порт.")
                return
            status_var.set(f"Открываю {port}…")

            def _on_code(code: str) -> None:
                # вызов из фонового потока — планируем в GUI
                self.root.after(0, lambda: _apply_code(code))

            def _on_err(msg: str) -> None:
                self.root.after(0, lambda: status_var.set(msg))

            sc = scanner_mod.SerialScanner(port, on_code=_on_code,
                                           on_error=_on_err)
            err = sc.open()
            if err:
                status_var.set(err)
                sc.close()
                return
            com_evt["scanner"] = sc
            status_var.set(f"Наблюдатель {port} открыт. Наведите сканер.")

        def _stop_watch() -> None:
            sc = com_evt.get("scanner")
            if sc is not None:
                try:
                    sc.close()
                except Exception:
                    pass
                com_evt["scanner"] = None
                status_var.set("Наблюдатель остановлен.")

        # ---------- возврат выбора в основное окно ----------
        def _apply_settings() -> None:
            mode, port = _selected_device()
            if not device_var.get():
                messagebox.showwarning("Сканер", "Не выбрано ни одно устройство.")
                return
            self.scanner_mode_var = mode
            self.scanner_port_var = port
            self._apply_scanner_enabled()
            if mode == "com":
                messagebox.showinfo("Сканер", f"Использую COM-порт {port}.")
            else:
                messagebox.showinfo("Сканер",
                                    "Использую сканер как клавиатуру (USB HID).")

        ttk.Label(win, textvariable=result_var, wraplength=430, justify="left"
                  ).pack(anchor="w", padx=16, pady=(4, 0))
        ttk.Label(win, textvariable=counter_var, foreground="#1a7f37"
                  ).pack(anchor="w", padx=16, pady=(2, 6))

        btn_row = ttk.Frame(win)
        btn_row.pack(pady=(4, 12))
        ttk.Button(btn_row, text="Обновить", command=_refresh_devs
                   ).pack(side="left", padx=4)
        ttk.Button(btn_row, text="Слушать COM",
                   command=_start_watch).pack(side="left", padx=4)
        ttk.Button(btn_row, text="Остановить",
                   command=_stop_watch).pack(side="left", padx=4)
        ttk.Button(btn_row, text="Применить",
                   command=_apply_settings).pack(side="left", padx=4)
        ttk.Button(btn_row, text="Закрыть",
                   command=win.destroy).pack(side="left", padx=4)

        def _on_close() -> None:
            _stop_watch()
            win.destroy()

        win.protocol("WM_DELETE_WINDOW", _on_close)
        entry.focus_set()

    def _render_code128_png(self, text: str, scale: int = 8) -> str | None:
        """Генерирует обычный Code128-штрихкод (не GS1) через zint и
        возвращает путь к PNG. text — строка, которую кодируем (например
        служебный код MERTECH '090400'). Возвращает None при ошибке."""
        from converter import find_zint_exe
        import subprocess, tempfile
        from pathlib import Path
        zint = find_zint_exe()
        if zint is None:
            return None
        td = tempfile.mkdtemp(prefix="dmtools_bc_")
        png = str(Path(td) / "code.png")
        try:
            proc = subprocess.run(
                [str(zint), "-b", "CODE128", "--scale", str(scale),
                 "-d", text, "-o", png],
                capture_output=True, text=True, timeout=20,
                creationflags=getattr(subprocess, "CREATE_NO_WINDOW", 0))
            if proc.returncode != 0 or not Path(png).is_file():
                return None
            return png
        except Exception:
            return None

    def _show_scanner_barcodes(self) -> None:
        """Окно со служебными штрихкодами MERTECH.

        Для сканеров MERTECH, подключённых как COM-порт, бывает нужно
        перевести сканер в режим USB-COM служебным штрихкодом. Здесь эти
        штрихкоды показаны на экране — сканер просто наводится на нужный
        код, как на бумажный.
        """
        from PIL import Image, ImageTk
        win = tk.Toplevel(self.root)
        win.title("Служебные штрихкоды MERTECH")
        win.resizable(False, False)
        win.transient(self.root)

        codes = [
            ("090400", "Перевести сканер в режим USB-COM",
             "Отсканируйте этот штрихкод сканером. После этого сканер "
             "будет передавать данные в COM-порт."),
            ("080404", "Очистить префиксы и суффиксы",
             "Отсканируйте также — чтобы сканер не добавлял к коду "
             "лишних символов."),
        ]

        tk.Label(win, text="Наведите сканер на нужный штрихкод, как на "
                           "бумажку.",
                 wraplength=480, justify="left", font=("Arial", 10)
                 ).pack(padx=16, pady=(12, 2))

        _refs = []

        for value, title, note in codes:
            png = self._render_code128_png(value)
            row = ttk.Frame(win)
            row.pack(fill="x", padx=16, pady=4)
            ttk.Label(row, text=title, font=("Arial", 10, "bold")
                      ).pack(anchor="w")
            if png:
                img = Image.open(png).convert("RGBA")
                photo = ImageTk.PhotoImage(img)
                _refs.append(photo)  # держим ссылку
                tk.Label(row, image=photo, bg="white"
                         ).pack(anchor="w", pady=2)
            ttk.Label(row, text=note + f"\nКод: {value}",
                      wraplength=480, justify="left", fg="#555"
                      ).pack(anchor="w")

        ttk.Button(win, text="Закрыть", command=win.destroy
                   ).pack(pady=(8, 12))

    def _show_scanner_diag(self) -> None:
        """Окно «Журнал диагностики сканера».

        Собирает автоматический отчёт о том, что Windows видит как
        устройства ввода и последовательные порты, а также логирует
        нажатия клавиш и байты с COM-порта в реальном времени.
        Журнал можно скопировать и переслать для разбора — в том числе
        «с другого компьютера», где подключён сканер.
        """
        import scanner_diag
        import scanner as scanner_mod

        win = tk.Toplevel(self.root)
        win.title("Журнал диагностики сканера")
        win.geometry("640x520")
        win.transient(self.root)

        log_text = scrolledtext.ScrolledText(
            win, wrap="none", font=("Consolas", 9), state="disabled")
        log_text.pack(fill="both", expand=True, padx=8, pady=(8, 4))

        def _append(line: str) -> None:
            log_text.configure(state="normal")
            log_text.insert("end", line + "\n")
            log_text.see("end")
            log_text.configure(state="disabled")

        # заголовок и начало журнала
        import time as _time
        _append("Содержание:")
        _append("  1) отчёт об оборудовании (COm/HID/USB с VID:PID);\n"
                "  2) живые события: нажатия клавиш и Enter от сканера;\n"
                "  3) если открыт COM-порт — принятые байты.")
        _append("")

        # строка-подсказка что делать
        hint = "Наведите сканер на код теперь — и нажмите Enter. "
        hint += "Байты/нажатия появятся в журнале."
        _append(">>> ПРОВЕРКА: " + hint)

        # --- живой захват нажатий (клавиатурный сканер) ---
        def _on_key(event=None) -> None:
            try:
                wname = event.widget.winfo_class() if event.widget else "?"
            except Exception:
                wname = "?"
            kc = getattr(event, "keysym", "?")
            # если это Enter/Return — особое сообщение
            if kc in ("Return", "KP_Enter"):
                _append(f"[клавиша] Enter нажата в поле {wname}")
            else:
                _append(f"[клавиша] {kc}")
        win.bind("<Key>", _on_key, add="+")

        # --- COM-наблюдатель (если выбран COM) ---
        diag_scanner = {"s": None}
        port = getattr(self, "scanner_port_var", None)
        if port and port.get():
            def _com_code(code: str) -> None:
                self.root.after(0, lambda: _append(f"[COM] код: {code}"))
            def _com_err(msg: str) -> None:
                self.root.after(0, lambda: _append(f"[COM] ошибка: {msg}"))
            try:
                sc = scanner_mod.SerialScanner(port.get(),
                                               on_code=_com_code,
                                               on_error=_com_err)
                err = sc.open()
                if err:
                    _append(f"[COM] {err}")
                else:
                    _append(f"[COM] слушаю {port.get()}…")
                    diag_scanner["s"] = sc
            except Exception as exc:
                _append(f"[COM] не удалось открыть: {exc}")

        # --- кнопка сбора отчёта об оборудовании ---
        def _collect_report() -> None:
            _append("")
            _append(">>> Отчёт об оборудовании…")
            report = scanner_diag.collect_diagnostic()
            for line in report.splitlines():
                _append(line)

        _append("")
        _append(">>> Нажмите «Собрать отчёт об оборудовании» ниже, а также")
        _append("    наведите сканер на такой код и посмотрите, что появится.")
        _append("")

        buttons = ttk.Frame(win)
        buttons.pack(fill="x", padx=8, pady=(0, 8))

        def _copy_all() -> None:
            txt = log_text.get("1.0", "end-1c")
            win.clipboard_clear()
            win.clipboard_append(txt)
            _append("[+] Журнал скопирован в буфер обмена")

        ttk.Button(buttons, text="Собрать отчёт об оборудовании",
                   command=_collect_report).pack(side="left", padx=3)
        ttk.Button(buttons, text="Копировать журнал",
                   command=_copy_all).pack(side="left", padx=3)
        ttk.Button(buttons, text="Закрыть", command=win.destroy
                   ).pack(side="left", padx=3)

        def _on_close() -> None:
            sc = diag_scanner.get("s")
            if sc is not None:
                try:
                    sc.close()
                except Exception:
                    pass
            win.destroy()
        win.protocol("WM_DELETE_WINDOW", _on_close)

    @staticmethod
    def _short_code(code: str) -> str:
        """Укорачивает код для показа, если он длинный."""
        return code if len(code) <= 40 else code[:37] + "…"

    @staticmethod
    def _looks_like_code(val: str) -> bool:
        """Эвристика: похож ли текст на код маркировки (длинная строка без
        пробелов, только цифры и допустимые символы GS-датаматрикса)."""
        if len(val) < 14:
            return False
        if any(ch.isspace() for ch in val):
            return False
        # допустимые символы: цифры, латиница, скобки (AI), разделитель GS, дефис
        allowed = set(
            "0123456789" + "ABCDEFGHIJKLMNOPQRSTUVWXYZabcdefghijklmnopqrstuvwxyz"
            + "()\x1d-#%&/.*"
        )
        return all(ch in allowed for ch in val)

    def _select_tab(self, tab_text: str) -> None:
        """Активирует (показывает и переключает) вкладку по её тексту."""
        try:
            self._tab_visible[tab_text].set(True)
            shown = [t for t in self._tab_order if self._tab_visible[t].get()]
            for tid in self.notebook.tabs():
                if self.notebook.tab(tid, "text") == tab_text:
                    self.notebook.select(tid)
                    return
        except Exception:
            pass

    def _set_quick_code(self, code: str) -> None:
        """Кладёт код в поле «Быстрый DataMatrix»."""
        if not hasattr(self, "quick_code_text"):
            return
        self.quick_code_text.delete("1.0", "end")
        self.quick_code_text.insert("1.0", code)
        self._tab_visible["Быстрый DataMatrix"].set(True)

    def _setup_tab_drag(self) -> None:
        """Включает перетаскивание вкладок мышкой (drag-and-drop).

        Нажатие на вкладку захватывает её; при движении курсора вкладка
        вставляется на новое место; отпускание фиксирует порядок.
        """
        if not hasattr(self, "notebook"):
            return
        self._drag_tab_index: int | None = None

        def _on_tab_press(event) -> None:
            try:
                idx = self.notebook.index(f"@{event.x},{event.y}")
                self._drag_tab_index = idx
            except Exception:
                self._drag_tab_index = None

        def _on_tab_motion(event) -> None:
            idx = self._drag_tab_index
            if idx is None:
                return
            try:
                target = self.notebook.index(f"@{event.x},{event.y}")
            except Exception:
                return
            if target is None or target == idx:
                return
            cur_id = self.notebook.tabs()[idx]
            self.notebook.insert(target, cur_id)
            self.notebook.select(target)
            self._drag_tab_index = target

        def _on_tab_release(_event) -> None:
            self._drag_tab_index = None

        self.notebook.bind("<Button-1>", _on_tab_press, add="+")
        self.notebook.bind("<B1-Motion>", _on_tab_motion, add="+")
        self.notebook.bind("<ButtonRelease-1>", _on_tab_release, add="+")

    def _preview_canvas_width(self) -> int:
        """Считывает ширину канваса полного предпросмотра В ГЛАВНОМ потоке
        (из фонового потока обращаться к Tcl нельзя). Возвращает реальную
        ширину или разумное значение по умолчанию."""
        try:
            if self.dm_full_preview_canvas.winfo_exists():
                self.root.update_idletasks()
                w = self.dm_full_preview_canvas.winfo_width()
                if w > 100:
                    return w
        except Exception:
            pass
        return 300

    def _preview_params(self) -> dict:
        """Собирает ВСЕ значения настроек этикетки в главном потоке, чтобы
        фоновый поток предпросмотра не обращался к Tk-переменным."""
        try:
            size_w, size_h = self._parse_label_size(self.dm_label_size_var.get())
        except ValueError:
            size_w, size_h = 58, 40
        return {
            "size_w": size_w,
            "size_h": size_h,
            "production_date": self.dm_date_var.get().strip(),
            "expiry": self.dm_expiry_var.get().strip(),
            "batch": self.dm_batch_var.get().strip(),
            "name": self.dm_name_var.get().strip(),
            "show_date": bool(self.dm_show_date_var.get()),
            "show_expiry": bool(self.dm_show_expiry_var.get()),
            "show_batch": bool(self.dm_show_batch_var.get()),
            "show_code": bool(self.dm_show_code_var.get()),
            "use_short": bool(self.dm_show_short_var.get()),
        }

    def _preview_dm(self) -> None:
        files = getattr(self, "_dm_csv_files", None) or []
        if not files:
            single = self.dm_csv_var.get().strip()
            if single:
                files = [Path(single)]
        if not files:
            messagebox.showerror("Ошибка", "Сначала выберите CSV-файл(ы)")
            return
        missing = [str(f) for f in files if not f.is_file()]
        if missing:
            messagebox.showerror("Ошибка", f"CSV-файл не найден:\n" + "\n".join(missing))
            return

        self.dm_status_var.set("Генерация предпросмотра...")

        # защита от одновременного запуска нескольких потоков (гонка):
        # если предпросмотр уже генерируется — не запускаем новый
        if getattr(self, "_preview_busy", False):
            return
        self._preview_busy = True

        # собираем настройки и ширину канваса ДО запуска фонового потока
        display_width = self._preview_canvas_width()
        params = self._preview_params()

        def work() -> None:
            try:
                csv_text = self._collect_dm_csv_text()
                combined, total = self._build_preview_images(
                    csv_text, max_codes=30, display_width=display_width, params=params
                )
                # отрисовка выполняется в главном потоке через очередь событий
                # (из фонового потока нельзя трогать Tk или вызывать root.after)
                self.events.put(("preview_draw", (combined, total)))
            except Exception as exc:
                self.events.put(("preview_error", str(exc)))

        threading.Thread(target=work, daemon=True).start()

    def _open_full_preview(self) -> None:
        """Открывает предпросмотр в отдельном модальном окне (Toplevel).
        Окно переиспользуется одним — если уже открыто, просто показываем его."""
        if getattr(self, "_preview_win", None) is not None and self._preview_win.winfo_exists():
            self._preview_win.deiconify()
            self._preview_win.lift()
            self._preview_win.focus_force()
            self._render_full_preview()
            return

        win = tk.Toplevel(self.root)
        win.title("Предпросмотр этикеток")
        win.transient(self.root)          # модальное поверх основного окна
        win.geometry("520x700")
        win.minsize(360, 480)
        win.protocol("WM_DELETE_WINDOW", lambda: self._close_preview())

        win.columnconfigure(0, weight=1)
        win.rowconfigure(1, weight=1)

        ttk.Label(
            win, text="Предпросмотр (до 30 этикеток). Нажмите «Обновить» "
                      "или измените настройки — перестроится по текущим параметрам."
        ).grid(row=0, column=0, sticky="w", padx=10, pady=(8, 4))

        # канвас предпросмотра в модальном окне
        self.dm_full_preview_canvas = tk.Canvas(
            win, bg="#ffffff", highlightthickness=0, yscrollcommand=None
        )
        self.dm_full_preview_canvas.grid(row=1, column=0, sticky="nsew", padx=(10, 0))
        self.dm_full_preview_scrollbar = ttk.Scrollbar(
            win, orient="vertical", command=self.dm_full_preview_canvas.yview
        )
        self.dm_full_preview_scrollbar.grid(row=1, column=1, sticky="ns", pady=(0, 0))
        self.dm_full_preview_canvas.configure(
            yscrollcommand=self.dm_full_preview_scrollbar.set
        )

        def _on_fullwheel(event) -> None:
            if self.dm_full_preview_canvas.winfo_exists():
                self.dm_full_preview_canvas.yview_scroll(
                    int(-1 * (event.delta / 120)), "units"
                )

        self.dm_full_preview_canvas.bind(
            "<Enter>", lambda _e: self.dm_full_preview_canvas.bind_all(
                "<MouseWheel>", _on_fullwheel)
        )
        self.dm_full_preview_canvas.bind(
            "<Leave>", lambda _e: self.dm_full_preview_canvas.unbind_all("<MouseWheel>")
        )

        btn_row = ttk.Frame(win)
        btn_row.grid(row=2, column=0, columnspan=2, sticky="ew", padx=10, pady=8)
        ttk.Button(
            btn_row, text="Обновить предпросмотр",
            command=lambda: self._render_full_preview(),
        ).pack(side="left")
        ttk.Button(
            btn_row, text="Закрыть", command=lambda: self._close_preview()
        ).pack(side="right")

        self._preview_win = win
        win.update_idletasks()
        self._render_full_preview()

    def _close_preview(self) -> None:
        """Закрывает модальное окно предпросмотра и освобождает ресурсы."""
        win = getattr(self, "_preview_win", None)
        if win is not None and win.winfo_exists():
            win.destroy()
        self._preview_win = None
        # сбрасываем изображение, чтобы не держать память
        self._dm_preview_images = []

    def _render_full_preview(self) -> None:
        """Запускает полный предпросмотр всех этикеток (в фоновом потоке).
        Результат рисуется в канвас модального окна `_preview_win`."""
        files = getattr(self, "_dm_csv_files", None) or []
        if not files:
            single = self.dm_csv_var.get().strip()
            if single:
                files = [Path(single)]
        if not files:
            messagebox.showerror("Ошибка", "Сначала выберите CSV-файл(ы)")
            return
        # если модальное окно закрыто — ничего не рисуем
        win = getattr(self, "_preview_win", None)
        if win is None or not win.winfo_exists():
            return
        self.dm_status_var.set("Генерация предпросмотра...")
        if getattr(self, "_preview_busy", False):
            return
        self._preview_busy = True

        # собираем настройки и ширину канваса ДО запуска фонового потока
        display_width = self._preview_canvas_width()
        params = self._preview_params()

        def work() -> None:
            try:
                csv_text = self._collect_dm_csv_text()
                combined, total = self._build_preview_images(
                    csv_text, max_codes=30, display_width=display_width, params=params
                )
                self.events.put(("preview_draw", (combined, total)))
            except Exception as exc:
                self.events.put(("preview_error", str(exc)))

        threading.Thread(target=work, daemon=True).start()

    def _build_preview_images(
        self,
        csv_text: str,
        *,
        max_codes: int,
        display_width: int,
        params: dict,
    ) -> tuple[object, int]:
        """Генерирует ОДНО общее изображение этикеток в фоновом потоке.

        Возвращает (комбинированное PIL-изображение, число этикеток).
        Компоновка в одно изображение позволяет показать результат одним
        кадром без мерцания канваса. НЕ трогает Tk-виджеты: все настройки
        и ширина канваса передаются параметрами, собранными в главном потоке.
        """
        import subprocess
        import tempfile
        from converter import (parse_csv_codes, gs1data_to_zint_data,
                               find_zint_exe, code_identification,
                               gs1data_to_zint_from_short)
        from PIL import Image, ImageDraw, ImageFont

        codes = parse_csv_codes(csv_text)
        codes = codes[:max_codes]
        zint_exe = find_zint_exe()
        if zint_exe is None:
            raise RuntimeError("Не найден zint.exe")

        size_w = params["size_w"]
        size_h = params["size_h"]
        production_date = params["production_date"]
        expiry = params["expiry"]
        batch = params["batch"]
        name = params["name"]
        show_date = params["show_date"]
        show_expiry = params["show_expiry"]
        show_batch = params["show_batch"]
        show_code = params["show_code"]
        use_short = bool(params.get("use_short", False))

        canvas_w = display_width
        if canvas_w < 100:
            canvas_w = 300
        max_display_w = canvas_w - 24
        disp_w = min(int(340), max_display_w)
        disp_h = int(round(disp_w * size_h / size_w))  # сохраняем пропорции (мм)
        if disp_h > 460:
            factor = 460 / disp_h
            disp_h = 460
            disp_w = int(disp_w * factor)
        scale = 1 + int(disp_w // 60)  # масштаб zint под размер кода

        try:
            font = ImageFont.truetype("arial.ttf", 11)
        except Exception:
            font = ImageFont.load_default()

        labels: list[Image.Image] = []
        for idx, code in enumerate(codes[:max_codes]):
            try:
                if use_short:
                    zint_data = gs1data_to_zint_from_short(code_identification(code))
                else:
                    zint_data = gs1data_to_zint_data(code)
                with tempfile.TemporaryDirectory() as td:
                    png_path = Path(td) / f"{idx}.png"
                    proc = subprocess.run(
                        [str(zint_exe), "-b", "DATAMATRIX", "--gs1", "--gssep",
                         "--square", f"--scale={scale}", "-d", zint_data,
                         "-o", str(png_path)],
                        capture_output=True, text=True,
                        creationflags=getattr(subprocess, "CREATE_NO_WINDOW", 0),
                    )
                    if proc.returncode != 0 or not png_path.exists():
                        continue

                    dm = Image.open(str(png_path)).convert("RGBA")
                    label = Image.new("RGBA", (disp_w, disp_h), "white")
                    draw = ImageDraw.Draw(label)

                    # DataMatrix — в верхней части этикетки (крупный)
                    m = 8
                    top_margin = 10  # небольшой отступ сверху
                    top_area_h = int(disp_h * 0.50)
                    dm_w = disp_w - 2 * m
                    dm_h = int(dm_w * dm.height / dm.width)
                    if dm_h > top_area_h - top_margin:
                        dm_h = top_area_h - top_margin
                        dm_w = int(dm_h * dm.width / dm.height)
                    dm_resized = dm.resize((dm_w, dm_h), Image.LANCZOS)
                    label.paste(dm_resized, ((disp_w - dm_w) // 2, top_margin), dm_resized)

                    # текст под штрихкодом — по центру этикетки (как в макете)
                    y_text = top_margin + dm_h + 6
                    lines = []
                    if show_date and production_date:
                        lines.append(f"Дата производства: {production_date}")
                    if show_expiry and expiry:
                        lines.append(f"Срок годности: {expiry}")
                    if show_batch and batch:
                        lines.append(f"Партия: {batch}")
                    if name:
                        lines.append(name)
                    if show_code:
                        lines.append(code.replace("\x1d", "\u29f8"))
                    for line in lines:
                        line_w = draw.textlength(line, font=font)
                        draw.text(((disp_w - line_w) / 2, y_text), line, fill="black", font=font)
                        y_text += 11
                    labels.append(label)
            except Exception:
                # ошибочный код (некорректный формат) — пропускаем, не срывая весь предпросмотр
                continue

        # компонуем все этикетки в одно вертикальное изображение
        if not labels:
            return Image.new("RGBA", (400, 100), "white"), 0
        gap = 8
        combined_h = sum(l.height for l in labels) + gap * (len(labels))
        combined = Image.new("RGBA", (max(l.width for l in labels) + 40, combined_h), "#f0f0f0")
        y_off = 14
        for i, lab in enumerate(labels):
            combined.paste(lab, (20, y_off), lab)
            y_off += lab.height + gap
        return combined, len(labels)

    def _draw_preview(self, combined: object, code_count: int) -> None:
        """Отрисовывает готовое комбинированное изображение этикеток одним
        кадром (главный поток). Устраняет мерцание канваса."""
        from PIL import ImageTk

        cv = self.dm_full_preview_canvas
        try:
            exists = cv.winfo_exists()
        except Exception:
            exists = False
        if not exists:
            # окно предпросмотра закрыли — канвас уничтожен, не рисуем
            self._preview_busy = False
            return
        cv.delete("all")
        cv.update_idletasks()  # убеждаемся, что канвас очищен до нового рисунка
        if combined is None:
            canvas_w = cv.winfo_width() or 300
            cv.configure(scrollregion=(0, 0, canvas_w, 100))
            self.dm_status_var.set(f"Показано этикеток: {code_count}")
            return
        photo = ImageTk.PhotoImage(combined)
        if not hasattr(self, "_dm_preview_images"):
            self._dm_preview_images = []
        self._dm_preview_images = [photo]
        cv.create_image(0, 0, image=photo, anchor="nw")
        canvas_w = cv.winfo_width() or 300
        total_h = combined.height + 20
        cv.configure(scrollregion=(0, 0, max(canvas_w, combined.width), total_h))
        self.dm_status_var.set(f"Показано этикеток: {code_count}")

    def _sizes_for_group(self, group: str) -> list[str]:
        """Доступные размеры этикеток в зависимости от выбранной группы."""
        all_sizes = ["25 × 25", "30 × 20", "40 × 30", "43 × 25", "50 × 30",
                     "57 × 40", "58 × 40", "58 × 60", "60 × 40", "90 × 60"]
        if group in ("Медизделия", "Косметика", "— без группы —"):
            return all_sizes
        if group in ("Духи/парфюмерия",):
            return ["25 × 25", "30 × 20", "43 × 25", "58 × 40"]
        if group == "Лёгкая промышленность":
            return ["40 × 30", "50 × 30", "58 × 40", "58 × 60", "60 × 40", "90 × 60"]
        if group == "Антисептики":
            return ["30 × 20", "43 × 25", "58 × 40"]
        return all_sizes

    def _on_show_change(self, _event=None) -> None:
        """При изменении галочек «что показывать» обновляет живой макет."""
        self._update_live_mock()

    def _update_live_mock(self) -> None:
        """Рисует схематичный макет одной этикетки по текущим настройкам.
        Мгновенно обновляется при любом изменении полей."""
        canvas = getattr(self, "dm_live_canvas", None)
        if canvas is None:
            return
        canvas.delete("all")
        w = canvas.winfo_width()
        if w < 50:
            w = 300
        h = canvas.winfo_height()
        if h < 50:
            h = 340

        try:
            size_w, size_h = self._parse_label_size(self.dm_label_size_var.get())
        except ValueError:
            size_w, size_h = 25, 25
        # масштаб: вписать этикетку в канвас с отступами
        scale = min((w - 40) / size_w, (h - 40) / size_h)
        lbl_w = size_w * scale
        lbl_h = size_h * scale
        x0 = (w - lbl_w) / 2
        y0 = 20

        # контур этикетки
        canvas.create_rectangle(
            x0, y0, x0 + lbl_w, y0 + lbl_h,
            outline="#555555", width=2, fill="#ffffff",
        )

        # реальный первый код из выбранного CSV (если есть)
        first_code = self._first_csv_code()
        dm_size = min(lbl_w * 0.62, lbl_h * 0.62)
        dm_x = x0 + (lbl_w - dm_size) / 2
        dm_y = y0 + 26

        if first_code:
            # Настоящий DataMatrix (по стандарту GS1/ECC200) подгружается
            # асинхронно через zint — с задержкой (debounce), чтобы не запускать
            # процесс на каждое нажатие клавиши при быстром вводе.
            def _launch_dm():
                self._insert_real_dm_async(
                    first_code, dm_x, dm_y, int(dm_size),
                    use_short=bool(self.dm_show_short_var.get()),
                )
            if getattr(self, "_live_dm_after_id", None):
                try:
                    self.root.after_cancel(self._live_dm_after_id)
                except Exception:
                    pass
            self._live_dm_after_id = self.root.after(200, _launch_dm)
            # пока картинка не готова — placeholder-рамка
            canvas.create_rectangle(
                dm_x, dm_y, dm_x + dm_size, dm_y + dm_size,
                outline="#cccccc", width=1, fill="#f8f8f8",
            )
            canvas.create_text(
                dm_x + dm_size / 2, dm_y + dm_size / 2,
                text="Загрузка...", fill="#999999", font=("Arial", 9),
            )
        else:
            # файл не загружен — место под код
            canvas.create_rectangle(
                dm_x, dm_y, dm_x + dm_size, dm_y + dm_size,
                outline="#cccccc", width=1, fill="#f8f8f8",
            )
            canvas.create_text(
                dm_x + dm_size / 2, dm_y + dm_size / 2,
                text="Загрузите файл", fill="#999999", font=("Arial", 9),
            )

        # текст под кодом: наименование, дата, партия, первый код
        tx = x0
        ty = dm_y + dm_size + 10
        mini = max(5, int(5 * (lbl_w / 150)))
        def line(s, bold=False):
            canvas.create_text(
                x0 + lbl_w / 2, ty, text=s or " ", fill="#333333",
                font=("Arial", mini, "bold" if bold else "normal"), anchor="n"
            )
        if self.dm_show_date_var.get() and self.dm_date_var.get().strip():
            line(f"Дата производства: {self.dm_date_var.get().strip()}", bold=self.dm_bold_date_var.get())
            ty += mini + 2
        if self.dm_show_expiry_var.get() and self.dm_expiry_var.get().strip():
            line(f"Срок годности: {self.dm_expiry_var.get().strip()}", bold=self.dm_bold_expiry_var.get())
            ty += mini + 2
        if self.dm_show_batch_var.get() and self.dm_batch_var.get().strip():
            line(f"Партия: {self.dm_batch_var.get().strip()}", bold=self.dm_bold_batch_var.get())
            ty += mini + 2
        if self.dm_name_var.get().strip():
            line(self.dm_name_var.get().strip(), bold=self.dm_bold_name_var.get())
            ty += mini + 2
        if self.dm_show_code_var.get() and first_code:
            shown = first_code.replace("\x1d", "␝")[:40]
            line(shown, bold=self.dm_bold_code_var.get())
            ty += mini + 2
        if not first_code:
            line("(файл не загружен)")
            ty += mini + 2

    def _insert_real_dm_async(self, code: str, x, y, size: int, use_short: bool = False) -> None:
        """Генерирует настоящий GS1 DataMatrix через zint в фоновом потоке
        (без окон) и вставляет картинку в макет по готовности.
        Только генерация PNG и resize выполняются в потоке; всё остальное,
        что связано с Tk (создание PhotoImage, отрисовка на канвасе), — в
        главном потоке через очередь событий."""
        canvas = getattr(self, "dm_live_canvas", None)
        if canvas is None:
            return

        def work() -> None:
            try:
                import subprocess
                import tempfile
                from pathlib import Path
                from converter import (
                    find_zint_exe, gs1data_to_zint_data,
                    code_identification, gs1data_to_zint_from_short,
                )
                from PIL import Image

                zint_exe = find_zint_exe()
                if zint_exe is None:
                    return
                if use_short:
                    data = gs1data_to_zint_from_short(code_identification(code))
                else:
                    data = gs1data_to_zint_data(code)
                with tempfile.TemporaryDirectory() as td:
                    png_path = Path(td) / "dm.png"
                    proc = subprocess.run(
                        [str(zint_exe), "-b", "DATAMATRIX", "--gs1", "--gssep",
                         "--square", f"--scale=8", "-d", data, "-o", str(png_path)],
                        capture_output=True, text=True,
                        creationflags=getattr(subprocess, "CREATE_NO_WINDOW", 0),
                    )
                    if proc.returncode != 0 or not png_path.exists():
                        return
                    img = Image.open(str(png_path)).convert("RGB")
                    img = img.resize((size, size), Image.LANCZOS)
                    # отдаём PIL-изображение в главный поток для отрисовки
                    self.events.put(("dm_live_insert", (img, x, y)))
            except Exception:
                pass

        threading.Thread(target=work, daemon=True).start()

    def _gs_char(self) -> str:
        return "\x1d"

    def _first_csv_code(self) -> str:
        """Возвращает первый реальный код из выбранных CSV-файлов (если есть).
        Использует кеш, вычисленный при выборе файлов, чтобы не перечитывать
        большой CSV при каждом обновлении живого макета."""
        cached = getattr(self, "_dm_cached_first_code", None)
        if cached is not None:
            return cached
        try:
            from converter import parse_csv_codes
            text = self._collect_dm_csv_text()
            codes = parse_csv_codes(text)
            self._dm_cached_first_code = codes[0] if codes else None
            return self._dm_cached_first_code or ""
        except Exception:
            return ""

        canvas.create_text(
            x0 + lbl_w / 2, y0 + lbl_h + 8,
            text=f"{size_w} × {size_h} мм",
            fill="#888888", font=("Arial", 9),
        )

    @staticmethod
    def _draw_dm_mock(canvas, x, y, nx, ny, cell) -> None:
        """Рисует прямоугольную сетку точек (схема DataMatrix)."""
        import random
        random.seed(42)
        for ix in range(nx):
            for iy in range(ny):
                fill = "#000000" if (ix + iy * 7) % 3 == 0 else "#ffffff"
                canvas.create_rectangle(
                    x + ix * cell, y + iy * cell,
                    x + (ix + 1) * cell, y + (iy + 1) * cell,
                    fill=fill, outline=fill,
                )

    def _on_preview_input_change(self, _event=None) -> None:
        """При изменении размера/даты/партии/наименования обновляет живой
        макет этикетки (мгновенно, без тяжёлой генерации zint)."""
        self._update_live_mock()

    def _format_date_input(self, text: str, date_format: str) -> str:
        """Приводит введённый текст к маске выбранного формата:
        расставляет точки по позициям. Допускает неполный ввод
        (точки ставятся только на заполненные позиции)."""
        digits = "".join(ch for ch in text if ch.isdigit())
        if date_format in ("ММ.ГГГГ", "ММ.ГГ"):
            digits = digits[:6] if date_format == "ММ.ГГГГ" else digits[:4]
            dots = (1,)
        elif date_format == "ДД.ММ.ГГ":
            digits = digits[:6]
            dots = (1, 3)
        else:  # ДД.ММ.ГГГГ
            digits = digits[:8]
            dots = (1, 3)
        formatted = ""
        for i, ch in enumerate(digits):
            formatted += ch
            if i in dots and i + 1 < len(digits):
                formatted += "."
        return formatted

    def _on_date_format(self, event=None) -> None:
        """Обрабатывает ввод в полях даты производства и срока годности:
        расставляет точки по маске выбранного формата, подсвечивает поле
        по валидности и обновляет живой макет этикетки."""
        if event is not None:
            widget = getattr(event, "widget", None)
            if widget == self.dm_expiry_entry:
                fmt = self.dm_expiry_format_var.get()
            elif widget == self.dm_expiry_fmt_combo:
                fmt = self.dm_expiry_format_var.get()
                widget = self.dm_expiry_entry
            elif widget == self.dm_date_fmt_combo:
                fmt = self.dm_date_format_var.get()
                widget = self.dm_date_entry
            else:
                widget = self.dm_date_entry
                fmt = self.dm_date_format_var.get()
        else:
            widget = self.dm_date_entry
            fmt = self.dm_date_format_var.get()
        if widget is None:
            return
        text = widget.get()
        formatted = self._format_date_input(text, fmt)
        if formatted != text:
            widget.delete(0, "end")
            widget.insert(0, formatted)
        widget.icursor(len(widget.get()))
        # подсветка поля по валидности
        if widget is self.dm_expiry_entry:
            self._mark_expiry_status()
        else:
            self._mark_date_status()
        # обновляем живой макет
        self._update_live_mock()

    def _mark_entry_status(self, widget, text: str, date_format: str,
                           allow_future: bool) -> str | None:
        """Подсвечивает поле: красная рамка при неверном значении.
        Возвращает текст ошибки (None — поле пустое или валидно)."""
        if widget is None:
            return None
        err = self._validate_date_text(text, date_format, allow_future=allow_future)
        try:
            if self.dark_var.get():
                bg = "#2b2b2b"
            else:
                bg = "#ffffff"
        except Exception:
            bg = "#ffffff"
        if err:
            widget.configure(background="#ffe0e0")
        else:
            widget.configure(background=bg)
        return err

    def _mark_date_status(self) -> None:
        """Подсвечивает поле даты производства по валидности и
        показывает под полем текст первой ошибки."""
        err = self._mark_entry_status(
            getattr(self, "dm_date_entry", None),
            self.dm_date_var.get().strip(),
            self.dm_date_format_var.get(),
            allow_future=False,
        )
        label = getattr(self, "dm_date_error_label", None)
        if label is not None:
            label.configure(text=err or "")

    def _mark_expiry_status(self) -> None:
        """Подсвечивает поле срока годности по валидности и
        показывает под полем текст первой ошибки.
        Будущее допускается (это конечная дата)."""
        err = self._mark_entry_status(
            getattr(self, "dm_expiry_entry", None),
            self.dm_expiry_var.get().strip(),
            self.dm_expiry_format_var.get(),
            allow_future=True,
        )
        label = getattr(self, "dm_expiry_error_label", None)
        if label is not None:
            label.configure(text=err or "")

    def _validate_date_text(self, text: str, date_format: str = "ММ.ГГГГ",
                            allow_future: bool = False):
        """Проверяет дату в одном из 4 форматов:
        ММ.ГГГГ / ММ.ГГ / ДД.ММ.ГГГГ / ДД.ММ.ГГ.
        Возвращает None, если дата валидна, или текст ошибки.
        allow_future=False — будущее запрещено (дата производства);
        allow_future=True — будущее допускается (срок годности).
        Для 2-значного года (ГГ) год интерпретируется как 2000+ГГ."""
        text = text.strip()
        if not text:
            return None  # поле пустое — не ошибка (дата опциональна)
        import datetime
        today = datetime.date.today()
        try:
            parts = text.split(".")
            full_year = date_format.endswith("ГГГГ")
            if len(parts) == 3:
                if date_format not in ("ДД.ММ.ГГГГ", "ДД.ММ.ГГ"):
                    return "Неверный формат даты: ожидается ММ.ГГГГ / ММ.ГГ"
                # ДД.ММ.ГГГГ или ДД.ММ.ГГ
                day, month, year = (int(p) for p in parts)
                if day < 1 or day > 31:
                    return f"Нет такого дня: {day}"
                if month < 1 or month > 12:
                    return f"Нет такого месяца: {month}"
                calendar_year = year + 2000 if not full_year else year
                if year < 0 or year > 99 and not full_year:
                    return "Год из двух цифр: 00..99"
                if full_year and year < 1900:
                    return "Год не может быть раньше 1900"
                import calendar
                max_day = calendar.monthrange(calendar_year, month)[1]
                if day > max_day:
                    return f"В месяце {month:02d} только {max_day} дней"
                try:
                    d = datetime.date(calendar_year, month, day)
                except ValueError:
                    return "Неверная дата"
                if not allow_future and d > today:
                    return "Дата производства не может быть позже сегодняшнего дня"
                return None
            elif len(parts) == 2:
                if date_format not in ("ММ.ГГГГ", "ММ.ГГ"):
                    return "Неверный формат даты: ожидается ДД.ММ.ГГГГ / ДД.ММ.ГГ"
                # ММ.ГГГГ или ММ.ГГ
                month, year = (int(p) for p in parts)
                if month < 1 or month > 12:
                    return f"Нет такого месяца: {month}"
                calendar_year = year + 2000 if not full_year else year
                if not full_year and (year < 0 or year > 99):
                    return "Год из двух цифр: 00..99"
                if full_year and year < 1900:
                    return "Год не может быть раньше 1900"
                if not allow_future and (calendar_year, month) > (today.year, today.month):
                    return "Дата производства не может быть позже текущего месяца"
                return None
            else:
                return (f"Неверный формат даты ({date_format})")
        except Exception:
            return "Неверная дата"

    def _invalid_date_reason(self) -> str | None:
        """Возвращает текст ошибки даты производства, если поле заполнено
        и невалидно, иначе None. Используется для блокировки генерации."""
        text = self.dm_date_var.get().strip()
        if self.dm_show_date_var.get() is False or not text:
            return None
        return self._validate_date_text(
            text, self.dm_date_format_var.get(), allow_future=False)

    def _invalid_expiry_reason(self) -> str | None:
        """Возвращает текст ошибки срока годности, если поле заполнено
        и невалидно, иначе None. Будущее всегда допускается."""
        text = self.dm_expiry_var.get().strip()
        if self.dm_show_expiry_var.get() is False or not text:
            return None
        return self._validate_date_text(
            text, self.dm_expiry_format_var.get(), allow_future=True)

    def _select_pdf(self) -> None:
        selected = filedialog.askopenfilenames(
            title="Выберите один или несколько PDF-файлов",
            filetypes=(("PDF files", "*.pdf"), ("All files", "*.*")),
        )
        if selected:
            paths = [Path(p) for p in selected]
            self._multi_source_files = paths
            if len(paths) == 1:
                path = paths[0]
                self.source_var.set(str(path))
                self.combined_output_var.set(str(default_combined_output_path(path)))
            else:
                self.source_var.set(
                    f"{len(paths)} файлов: "
                    + "; ".join(str(p) for p in paths[:3])
                    + (" ..." if len(paths) > 3 else "")
                )
                self.combined_output_var.set(
                    str(default_combined_output_path(paths[0].parent))
                )
            self._update_output_controls()

    def _select_source_folder(self) -> None:
        selected = filedialog.askdirectory(title="Выберите папку с PDF-файлами")
        if selected:
            self._multi_source_files = []
            path = Path(selected)
            self.source_var.set(str(path))
            self.combined_output_var.set(str(default_combined_output_path(path)))

    def _select_output_file(self) -> None:
        current_text = self.combined_output_var.get().strip()
        current_path = Path(current_text) if current_text else None
        selected = filedialog.asksaveasfilename(
            title="Выберите общий CSV-файл",
            initialdir=str(current_path.parent) if current_path else None,
            initialfile=current_path.name if current_path else None,
            defaultextension=".csv",
            filetypes=(("CSV files", "*.csv"), ("All files", "*.*")),
            confirmoverwrite=False,
        )
        if selected:
            self.combined_output_var.set(selected)

    def _open_output_folder(self) -> None:
        target = self.last_output_dir
        if target and Path(target).is_dir():
            try:
                os.startfile(target)  # type: ignore[attr-defined]
                return
            except OSError:
                pass
        messagebox.showinfo(
            "Папка не найдена",
            "Сначала выполните конвертацию, чтобы было что открыть.",
        )

    def _export_excel(self) -> None:
        codes = list(self._last_codes)
        if not codes:
            messagebox.showinfo(
                "Нет данных",
                "Сначала выполните конвертацию, чтобы появились коды.",
            )
            return
        default_dir = self.last_output_dir or str(Path.home())
        selected = filedialog.asksaveasfilename(
            title="Сохранить Excel-файл",
            initialdir=default_dir,
            initialfile="Коды_маркировки.xlsx",
            defaultextension=".xlsx",
            filetypes=(("Excel files", "*.xlsx"), ("All files", "*.*")),
            confirmoverwrite=False,
        )
        if not selected:
            return
        path = Path(selected)
        try:
            write_codes_xlsx(path, codes, escape_mode=self.escape_mode_var.get())
        except Exception as exc:
            messagebox.showerror("Ошибка", f"Не удалось сохранить Excel-файл:\n{exc}")
            return
        self.last_output_dir = str(path.parent)
        self.status_var.set("Excel-файл сохранён")
        self._append_log(f"Excel сохранён: {path}\n  Листов: 10 (все варианты)\n\n")
        messagebox.showinfo("Готово", f"Excel-файл сохранён:\n{path}")

    def _open_codes_dialog(self) -> None:
        codes = list(self._last_codes)
        if not codes:
            messagebox.showinfo(
                "Нет данных",
                "Сначала выполните конвертацию, чтобы появились коды.",
            )
            return

        def display_mc(mc: MarkingCode) -> str:
            return apply_escape_mode(full_code_str(mc), self.escape_mode_var.get())

        dialog = tk.Toplevel(self.root)
        dialog.title("Открыть код")
        dialog.geometry("860x560")
        dialog.minsize(620, 380)
        dialog.transient(self.root)
        dialog.grab_set()

        dialog.update_idletasks()
        main_x = self.root.winfo_rootx()
        main_y = self.root.winfo_rooty()
        main_w = self.root.winfo_width()
        main_h = self.root.winfo_height()
        dlg_w = dialog.winfo_width()
        dlg_h = dialog.winfo_height()
        x = main_x + max(0, (main_w - dlg_w) // 2)
        y = main_y + max(0, (main_h - dlg_h) // 2)
        dialog.geometry(f"+{x}+{y}")

        dark = self.dark_var.get()
        if dark:
            dlgbg = "#1f1f1f"
            fg = "#e8e8e8"
        else:
            dlgbg = "#f0f0f0"
            fg = "#000000"
        dialog.configure(bg=dlgbg)

        wrapper = ttk.Frame(dialog)
        wrapper.pack(fill="both", expand=True, padx=10, pady=(10, 4))

        seen: set[str] = set()
        dup_raw: set[str] = set()
        invalid_gtin = 0
        gtin_counts: dict[str, int] = {}
        for mc in codes:
            raw_str = full_code_str(mc)
            if raw_str in seen:
                dup_raw.add(raw_str)
            else:
                seen.add(raw_str)
            if not gtin_check_digit_is_valid(mc.gtin):
                invalid_gtin += 1
            gtin_counts[mc.gtin] = gtin_counts.get(mc.gtin, 0) + 1
        summary_text = (
            f"Кодов: {len(codes)}"
            f"  ·  уникальных: {len(seen)}"
            f"  ·  ошибок GTIN: {invalid_gtin}"
            f"  ·  дублей: {len(dup_raw)}"
        )
        ttk.Label(wrapper, text=summary_text, style="TLabel").pack(fill="x", pady=(0, 6))

        canvas = tk.Canvas(wrapper, highlightthickness=0, bg=dlgbg, borderwidth=0)
        scrollbar = ttk.Scrollbar(wrapper, orient="vertical", command=canvas.yview)
        list_frame = ttk.Frame(canvas, style="TFrame")
        list_window = canvas.create_window((0, 0), window=list_frame, anchor="nw")

        list_frame.bind(
            "<Configure>", lambda _e: canvas.configure(scrollregion=canvas.bbox("all"))
        )
        canvas.configure(yscrollcommand=scrollbar.set)
        canvas.pack(side="left", fill="both", expand=True)
        scrollbar.pack(side="right", fill="y")

        def copy_single(mc: MarkingCode) -> None:
            self.root.clipboard_clear()
            self.root.clipboard_append(display_mc(mc))
            self.status_var.set("Код скопирован")

        for mc in codes:
            row = ttk.Frame(list_frame, style="TFrame")
            row.pack(fill="x", pady=1)

            raw_str = full_code_str(mc)
            is_valid = gtin_check_digit_is_valid(mc.gtin)
            is_dup = raw_str in dup_raw
            gtin_qty = gtin_counts.get(mc.gtin, 0)

            dot_color = "#2CA847" if is_valid else "#d33"
            dot = tk.Canvas(row, width=12, height=12, bg=dlgbg, highlightthickness=0, bd=0)
            dot.create_oval(2, 2, 11, 11, fill=dot_color, outline="")
            dot.pack(side="left", padx=(0, 6))

            code_label = tk.Label(
                row, text=display_mc(mc), font=("Consolas", 10),
                bg=dlgbg, fg=fg, anchor="w",
            )
            code_label.pack(side="left", fill="x", expand=True)

            tk.Label(row, text=f"× {gtin_qty}", font=("Segoe UI", 9),
                     bg=dlgbg, fg="#888888", anchor="e").pack(side="left", padx=(6, 0))

            status_text = "Ошибка GTIN" if not is_valid else ("Дубль" if is_dup else "OK")
            status_color = ("#b8860b" if is_dup else ("#d33" if not is_valid else "#2CA847"))
            tk.Label(row, text=status_text, font=("Segoe UI", 9),
                     bg=dlgbg, fg=status_color, anchor="e").pack(side="left", padx=(8, 0))
            ttk.Button(
                row, text="Копировать", command=lambda m=mc: copy_single(m)
            ).pack(side="right", padx=(10, 0))

        buttons = ttk.Frame(dialog)
        buttons.pack(fill="x", padx=10, pady=10)

        def copy_all() -> None:
            data = "\n".join(display_mc(mc) for mc in codes)
            self.root.clipboard_clear()
            self.root.clipboard_append(data)
            self.status_var.set("Все коды скопированы")
            messagebox.showinfo("Готово", "Все коды скопированы в буфер обмена")

        ttk.Button(buttons, text="Скопировать все", command=copy_all).pack(side="left")
        ttk.Button(buttons, text="Закрыть", command=dialog.destroy).pack(side="left", padx=(8, 0))
        dialog.after(10, lambda: dialog.lift())

    def _update_output_controls(self) -> None:
        merge_csv = self.merge_csv_var.get()
        if merge_csv and not self.combined_output_var.get().strip():
            source_text = self.source_var.get().strip()
            if source_text:
                source = Path(source_text).expanduser()
                if source.exists():
                    self.combined_output_var.set(str(default_combined_output_path(source)))

        enabled = not self.running and merge_csv
        state = "normal" if enabled else "disabled"
        self.output_entry.configure(state=state)
        self.select_output_button.configure(state=state)

    def _start(self) -> None:
        if self.running:
            return

        source_text = self.source_var.get().strip()
        output_text = self.combined_output_var.get().strip()
        merge_csv = self.merge_csv_var.get()
        escape_mode = self.escape_mode_var.get()
        if not source_text:
            messagebox.showerror("Ошибка", "Не выбран PDF-файл или папка")
            return

        source = Path(source_text).expanduser()
        if not self._multi_source_files and not source.exists():
            messagebox.showerror("Ошибка", f"Источник не найден:\n{source}")
            return

        combined_output_path: Path | None = None
        if merge_csv:
            if not output_text:
                messagebox.showerror("Ошибка", "Не указан путь к общему CSV-файлу")
                return
            combined_output_path = Path(output_text).expanduser()
            if combined_output_path.suffix.lower() != ".csv":
                messagebox.showerror("Ошибка", "Путь к общему файлу должен иметь расширение .csv")
                return
            if combined_output_path.exists():
                if not combined_output_path.is_file():
                    messagebox.showerror("Ошибка", f"Путь к общему CSV не является файлом:\n{combined_output_path}")
                    return
                if not messagebox.askyesno(
                    "Перезапись CSV",
                    f"Файл уже существует:\n{combined_output_path}\n\nПерезаписать его?",
                ):
                    return
            try:
                combined_output_path.parent.mkdir(parents=True, exist_ok=True)
            except OSError as exc:
                messagebox.showerror("Ошибка", f"Не удалось создать папку для общего CSV:\n{exc}")
                return

        self._clear_log()
        self.cancel_event.clear()
        self.running = True
        self._set_controls_running(True)
        self.progress.configure(value=0, maximum=1)
        self.status_var.set("Поиск PDF-файлов...")

        if self._multi_source_files:
            source_files = list(self._multi_source_files)
        else:
            source_files = find_pdf_files(source)

        self.worker = threading.Thread(
            target=self._run_batch,
            args=(
                source,
                source_files,
                combined_output_path.resolve() if combined_output_path is not None else None,
                escape_mode,
            ),
            daemon=True,
        )
        self.worker.start()

    def _run_batch(self, source: Path, source_files: list[Path],
                   combined_output_path: Path | None, escape_mode: str) -> None:
        pdf_files = source_files
        if not pdf_files:
            self.events.put(("log", f"PDF-файлы не найдены: {source}\n"))
            self.events.put(("done", (0, 0, False)))
            return

        jobs = allocate_output_paths(pdf_files)
        self.events.put(("maximum", len(jobs)))
        self.events.put(("log", f"Найдено PDF-файлов: {len(jobs)}\n"))
        if combined_output_path is None:
            self.events.put(("log", "CSV сохраняются рядом с исходными PDF.\n\n"))
        else:
            self.events.put(("log", f"Общий CSV: {combined_output_path}\n\n"))
        mode_names = {
            "gs": "Разделитель GS: символ GS (обычный)",
            "1c": "Разделитель GS: для 1С (замена на \\x1d)",
            "ozon": "Разделитель GS: для OZON (замена на \\u001d)",
        }
        self.events.put(("log", mode_names.get(escape_mode, "Разделитель GS") + "\n\n"))
        self.events.put(("log", "Вывод: полный код маркировки\n\n"))

        succeeded = 0
        failed = 0
        cancelled = False
        combined_codes: list[MarkingCode] = []
        all_codes: list[MarkingCode] = []

        for index, (pdf_path, csv_path) in enumerate(jobs, start=1):
            if self.cancel_event.is_set():
                cancelled = True
                break

            self.events.put(
                (
                    "log",
                    f"{'=' * 72}\n[{index}/{len(jobs)}] {pdf_path}\n"
                    + (f"CSV: {csv_path.name}\n" if combined_output_path is None else ""),
                )
            )
            self.events.put(("status", f"Файл {index} из {len(jobs)}"))

            def page_progress(current_page: int, total_pages: int,
                              file_index: int = index, total_files: int = len(jobs)) -> None:
                self.events.put(
                    ("status", f"Файл {file_index}/{total_files}, страница {current_page}/{total_pages}")
                )

            try:
                decoded = decode_pdf(
                    pdf_path,
                    log=lambda message: self.events.put(("log", message)),
                    page_progress=page_progress,
                    cancel_event=self.cancel_event,
                )
                all_codes.extend(decoded.codes)
                if combined_output_path is None:
                    write_codes_csv(
                        csv_path, list(decoded.codes),
                        escape_csv=False, escape_mode=escape_mode,
                    )
                else:
                    combined_codes.extend(decoded.codes)
                self.events.put(
                    (
                        "log",
                        f"Decoded: {decoded.code_count}\n"
                        f"Unique codes: {decoded.unique_code_count}\n"
                        f"GTIN counts: {decoded.gtins}\n",
                    )
                )
            except ConversionCancelled:
                cancelled = True
                self.events.put(("log", "\nКонвертация остановлена пользователем.\n"))
                break
            except ConversionError as exc:
                failed += 1
                self.events.put(("log", f"ERROR: {exc}\nРезультат: ошибка\n\n"))
            except Exception as exc:
                failed += 1
                self.events.put(("log", f"UNEXPECTED ERROR: {exc}\nРезультат: ошибка\n\n"))
            else:
                succeeded += 1
                self.events.put(("log", "Результат: успешно\n\n"))
                if combined_output_path is None:
                    self.last_output_dir = str(csv_path.parent)
                self._log_op(
                    f"Декодирован PDF: {pdf_path} -> {csv_path} "
                    f"(кодов: {decoded.code_count}, уникальных: {decoded.unique_code_count})"
                )

            self.events.put(("progress", index))

        if combined_output_path is not None and not cancelled and succeeded:
            try:
                write_codes_csv(
                    combined_output_path, combined_codes,
                    escape_csv=False, escape_mode=escape_mode,
                )
            except Exception as exc:
                failed += succeeded
                succeeded = 0
                self.events.put(("log", f"ERROR: не удалось записать общий CSV: {exc}\nРезультат: ошибка\n"))
            else:
                self.last_output_dir = str(combined_output_path.parent)
                self.events.put(
                    ("log", f"Общий CSV записан: {combined_output_path}\nВсего кодов: {len(combined_codes)}\n")
                )
        elif combined_output_path is not None and cancelled:
            self.events.put(("log", "Общий CSV не записан из-за остановки.\n"))

        self._last_codes = all_codes
        self.events.put(("done", (succeeded, failed, cancelled)))

    def _stop(self) -> None:
        if not self.running:
            return
        self.cancel_event.set()
        self.status_var.set("Остановка...")
        self.stop_button.configure(state="disabled")

    def _drain_events(self) -> None:
        try:
            while True:
                event_type, payload = self.events.get_nowait()
                if event_type == "log":
                    self._append_log(str(payload))
                    self._global_log(str(payload).rstrip())
                elif event_type == "maximum":
                    self.progress.configure(maximum=int(payload), value=0)
                elif event_type == "progress":
                    self.progress.configure(value=int(payload))
                elif event_type == "status":
                    self.status_var.set(str(payload))
                elif event_type == "done":
                    succeeded, failed, cancelled = payload  # type: ignore[misc]
                    self._finish(int(succeeded), int(failed), bool(cancelled))
                elif event_type == "preview_draw":
                    combined, total = payload  # type: ignore[misc]
                    # отрисовываем в главном потоке и снимаем блокировку
                    self._preview_busy = False
                    try:
                        self._draw_preview(combined, total)
                    except Exception as exc:  # pragma: no cover
                        messagebox.showerror("Ошибка", str(exc))
                        self.dm_status_var.set("Ошибка")
                elif event_type == "preview_error":
                    # отрисовка завершилась ошибкой — снимаем блокировку и сообщаем
                    self._preview_busy = False
                    self._global_log(f"Ошибка предпросмотра: {payload}")
                    messagebox.showerror("Ошибка", str(payload))
                    self.dm_status_var.set("Ошибка")
                elif event_type == "dm_generated":
                    self._handle_dm_generated(payload)  # type: ignore[arg-type]
                elif event_type == "dm_printed":
                    self._handle_dm_printed(payload)  # type: ignore[arg-type]
                elif event_type == "dm_png_exported":
                    self._handle_dm_png_exported(payload)  # type: ignore[arg-type]
                elif event_type == "dm_live_insert":
                    self._handle_dm_live_insert(payload)  # type: ignore[arg-type]
        except queue.Empty:
            pass

        self.root.after(100, self._drain_events)

    def _handle_dm_generated(self, data: dict) -> None:
        """Обрабатывает результат генерации PDF (запускается в ГЛАВНОМ потоке).
        Раскрывает все GUI-обновления, журнал и авто-открытие свежего файла,
        которые нельзя было делать из фонового потока генерации."""
        pages = int(data.get("pages", 0))
        err = str(data.get("err", "") or "")
        actual = data.get("save_path", "")
        out_path = data.get("out_path", "")
        total_codes = int(data.get("total_codes", 0))
        diag = data.get("diag")

        self._generating = False
        self._close_progress_dialog()
        self.dm_progress.configure(value=100)
        self.dm_progress_label.configure(text="Готово")
        self.dm_generate_button.configure(state="normal")

        if diag is not None:
            self._append_log(f"[diag] {diag}\n")

        if pages > 0:
            self._last_generated_pdf = actual
            # сразу открываем свежий PDF в главном потоке (root.after здесь
            # допустим — мы в главном потоке)
            self.root.after(200, lambda p=actual: _safe_startfile(str(p)))

        if err and pages == 0:
            messagebox.showerror("Ошибка", err)
            self.dm_status_var.set("Ошибка")
            self._log_op(f"Генерация PDF не удалась: {err}")
            return

        self.last_output_dir = str(Path(out_path).parent) if out_path else self.last_output_dir
        self._log_op(
            f"Сгенерирован PDF: {actual or out_path} "
            f"(страниц: {pages} из {total_codes})"
        )
        if err:
            self.dm_status_var.set(f"Готово: {pages} из {total_codes} (часть пропущена)")
            self._append_log(f"Создан PDF: {out_path}\n  Страниц: {pages}\n")
            # предупреждение про копию/пропуск — только в журнал, без окна
            self._append_log(err + "\n")
        else:
            self.dm_status_var.set(f"Готово: {pages} из {total_codes}")
            self._append_log(f"Создан PDF: {actual}\n  Страниц: {pages}\n")

    def _handle_dm_printed(self, data: dict) -> None:
        """Обрабатывает результат подготовки PDF к печати (главный поток)."""
        from converter import print_pdf_windows

        pages = int(data.get("pages", 0))
        err = str(data.get("err", "") or "")
        tmp_pdf = data.get("tmp_pdf", "")

        self.dm_print_button.configure(state="normal")

        if err and pages == 0:
            messagebox.showerror("Ошибка", err)
            return

        if err:
            messagebox.showwarning("Предупреждение", err)

        perr = print_pdf_windows(Path(tmp_pdf))
        if perr:
            messagebox.showerror("Ошибка", perr)
            self._log_op(f"Печать не запущена: {perr}")
        else:
            self._log_op(f"Отправлено на печать: {pages} этикеток")
            self.dm_status_var.set(f"Печать открыта: {pages} этикеток")

    def _handle_dm_png_exported(self, data: dict) -> None:
        """Обрабатывает результат экспорта этикеток в PNG (главный поток)."""
        saved = int(data.get("saved", 0))
        err = str(data.get("err", "") or "")
        total_codes = int(data.get("total_codes", 0))
        out_dir = str(data.get("out_dir", ""))

        self.dm_exportpng_button.configure(state="normal")
        self._generating = False
        self._close_progress_dialog()

        if err:
            messagebox.showerror("Ошибка", err)
            self._log_op(f"Экспорт PNG не удался: {err}")
        else:
            self._log_op(f"Экспорт PNG: сохранено {saved} из {total_codes} в {out_dir}")
            messagebox.showinfo(
                "Экспорт завершён",
                f"Сохранено этикеток: {saved} из {total_codes}\nв папку:\n{out_dir}",
            )
            self.dm_status_var.set(f"Экспорт: {saved} из {total_codes}")

    def _handle_dm_live_insert(self, payload) -> None:
        """Вставляет готовый DataMatrix в живой макет (главный поток).
        PhotoImage создаётся здесь, в главном потоке — так допустимо."""
        from PIL import ImageTk

        img, x, y = payload
        canvas = getattr(self, "dm_live_canvas", None)
        if canvas is None or img is None:
            return
        try:
            photo = ImageTk.PhotoImage(img)
            if not hasattr(self, "_live_photo"):
                self._live_photo = None
            self._live_photo = photo
            canvas.create_image(int(x), int(y), image=photo, anchor="nw")
        except Exception:
            pass

    def _finish(self, succeeded: int, failed: int, cancelled: bool) -> None:
        self.running = False
        self._set_controls_running(False)
        if cancelled:
            status = f"Остановлено. Успешно: {succeeded}, ошибок: {failed}"
        else:
            status = f"Готово. Успешно: {succeeded}, ошибок: {failed}"
        self.status_var.set(status)
        self._append_log(f"\n{status}\n")

    def _set_controls_running(self, running: bool) -> None:
        normal_state = "disabled" if running else "normal"
        for widget in (
            self.source_entry,
            self.select_file_button,
            self.select_folder_button,
            self.merge_csv_check,
            self.escape_gs_radio,
            self.escape_1c_radio,
            self.escape_ozon_radio,
            self.start_button,
            self.excel_button,
        ):
            widget.configure(state=normal_state)
        self._update_output_controls()
        self.stop_button.configure(state="normal" if running else "disabled")

    def _append_log(self, text: str) -> None:
        # внутренний журнал вкладки убран — всё идёт в общий «Журнал»
        self._global_log(str(text).rstrip())

    def _clear_log(self) -> None:
        # внутренний журнал вкладки убран — очищаем общий «Журнал»
        self._clear_global_log()

    def _log_op(self, message: str) -> None:
        """Пишет строку операции с меткой времени в файл журнала и в
        отдельную вкладку «Журнал» (общий лог событий программы)."""
        import datetime
        try:
            stamp = datetime.datetime.now().strftime("%Y-%m-%d %H:%M:%S")
            log_file = log_file_path()
            line = f"[{stamp}] {message}\n"
            with open(log_file, "a", encoding="utf-8") as fh:
                fh.write(line)
        except OSError:
            pass
        self._global_log(f"[{stamp}] {message}")

    def _global_log(self, message: str) -> None:
        """Пишет строку в отдельную вкладку «Журнал» (общий лог событий).
        Безопасен: вызывается из главного потока (Tk-виджет)."""
        log = getattr(self, "global_log", None)
        if log is None:
            return
        try:
            log.configure(state="normal")
            log.insert("end", message + "\n")
            log.see("end")   # автоскролл вниз
            log.configure(state="disabled")
        except tk.TclError:
            pass

    def _clear_global_log(self) -> None:
        """Очищает содержимое вкладки «Журнал»."""
        log = getattr(self, "global_log", None)
        if log is None:
            return
        try:
            log.configure(state="normal")
            log.delete("1.0", "end")
            log.configure(state="disabled")
        except tk.TclError:
            pass

    def _refresh_profiles_combo(self) -> None:
        """Обновляет список шаблонов в выпадающем списке."""
        if not hasattr(self, "dm_profiles_combo"):
            return
        names = list(self._profiles.keys())
        self.dm_profiles_combo.configure(values=names)
        if self.dm_profile_name_var.get() not in names:
            self.dm_profile_name_var.set("")

    def _collect_profile(self) -> dict:
        """Собирает текущие настройки этикетки в dict для шаблона."""
        return {
            "label_size": self.dm_label_size_var.get(),
            "group": self.dm_group_var.get(),
            "show_date": bool(self.dm_show_date_var.get()),
            "show_expiry": bool(self.dm_show_expiry_var.get()),
            "show_batch": bool(self.dm_show_batch_var.get()),
            "show_code": bool(self.dm_show_code_var.get()),
            "show_short": bool(self.dm_show_short_var.get()),
            "name": self.dm_name_var.get(),
            "date": self.dm_date_var.get(),
            "expiry": self.dm_expiry_var.get(),
            "date_format": self.dm_date_format_var.get(),
            "expiry_format": self.dm_expiry_format_var.get(),
            "batch": self.dm_batch_var.get(),
            "bold_name": bool(self.dm_bold_name_var.get()),
            "bold_date": bool(self.dm_bold_date_var.get()),
            "bold_expiry": bool(self.dm_bold_expiry_var.get()),
            "bold_batch": bool(self.dm_bold_batch_var.get()),
            "bold_code": bool(self.dm_bold_code_var.get()),
            "size_name": self.dm_size_name_var.get(),
            "size_date": self.dm_size_date_var.get(),
            "size_expiry": self.dm_size_expiry_var.get(),
            "size_batch": self.dm_size_batch_var.get(),
            "size_code": self.dm_size_code_var.get(),
        }

    def _apply_profile(self, profile: dict) -> None:
        """Применяет словарь настроек шаблона к полям формы."""
        if not isinstance(profile, dict):
            return
        mapping = {
            "label_size": (self.dm_label_size_var, str),
            "group": (self.dm_group_var, str),
            "name": (self.dm_name_var, str),
            "date": (self.dm_date_var, str),
            "expiry": (self.dm_expiry_var, str),
            "batch": (self.dm_batch_var, str),
            "size_name": (self.dm_size_name_var, str),
            "size_date": (self.dm_size_date_var, str),
            "size_expiry": (self.dm_size_expiry_var, str),
            "size_batch": (self.dm_size_batch_var, str),
            "size_code": (self.dm_size_code_var, str),
            "date_format": (self.dm_date_format_var, str),
            "expiry_format": (self.dm_expiry_format_var, str),
        }
        for key, (var, caster) in mapping.items():
            if key in profile and isinstance(profile[key], str):
                try:
                    var.set(caster(profile[key]))
                except Exception:
                    pass
        bool_map = {
            "show_date": self.dm_show_date_var,
            "show_expiry": self.dm_show_expiry_var,
            "show_batch": self.dm_show_batch_var,
            "show_code": self.dm_show_code_var,
            "show_short": self.dm_show_short_var,
            "bold_name": self.dm_bold_name_var,
            "bold_date": self.dm_bold_date_var,
            "bold_expiry": self.dm_bold_expiry_var,
            "bold_batch": self.dm_bold_batch_var,
            "bold_code": self.dm_bold_code_var,
        }
        for key, var in bool_map.items():
            if key in profile:
                try:
                    var.set(bool(profile[key]))
                except Exception:
                    pass
        # корректируем список размеров под группу
        group = self.dm_group_var.get()
        sizes = self._sizes_for_group(group)
        if hasattr(self, "dm_size_combo"):
            self.dm_size_combo.configure(values=sizes)
        if self.dm_label_size_var.get() not in sizes:
            self.dm_label_size_var.set(sizes[0] if sizes else "58 × 40")
        self._update_live_mock()

    def _apply_selected_profile(self, _event=None) -> None:
        """Применяет шаблон, выбранный из выпадающего списка."""
        name = self.dm_profile_name_var.get()
        profile = self._profiles.get(name)
        if profile:
            self._apply_profile(profile)
            self.dm_status_var.set(f"Загружен шаблон: {name}")
            self._log_op(f"Загружен шаблон этикеток: {name}")

    def _save_profile_dialog(self) -> None:
        """Сохраняет текущие настройки как шаблон этикетки."""
        if not hasattr(self, "dm_profiles_combo"):
            return
        name = self.dm_profile_name_var.get().strip()
        if not name:
            from tkinter import simpledialog

            name = simpledialog.askstring(
                "Сохранить шаблон",
                "Введите имя шаблона:",
                parent=self.root,
            )
            if not name:
                return
            name = name.strip()
            if not name:
                return
        self._profiles[name] = self._collect_profile()
        self._refresh_profiles_combo()
        self.dm_profile_name_var.set(name)
        self.dm_status_var.set(f"Шаблон сохранён: {name}")
        self._log_op(f"Сохранён шаблон этикеток: {name}")
        self._save_config()

    def _load_profile_dialog(self) -> None:
        """Загружает сохранённый шаблон этикетки."""
        if not hasattr(self, "dm_profiles_combo"):
            return
        names = list(self._profiles.keys())
        if not names:
            messagebox.showinfo("Шаблоны", "Сохранённых шаблонов пока нет.")
            return
        from tkinter import simpledialog

        picked = simpledialog.askstring(
            "Загрузить шаблон",
            "Введите имя шаблона:\n\n" + "\n".join(names),
            parent=self.root,
        )
        if not picked:
            return
        picked = picked.strip()
        profile = self._profiles.get(picked)
        if profile is None:
            # пробуем нечёткое совпадение по подстроке
            for n in names:
                if picked.lower() in n.lower():
                    profile = self._profiles[n]
                    picked = n
                    break
        if profile is None:
            messagebox.showerror("Ошибка", f"Шаблон не найден:\n{picked}")
            return
        self._apply_profile(profile)
        self.dm_profile_name_var.set(picked)
        self.dm_status_var.set(f"Загружен шаблон: {picked}")
        self._log_op(f"Загружен шаблон этикеток: {picked}")

    def _load_config(self) -> None:
        cfg = config_path()
        try:
            with open(cfg, "r", encoding="utf-8") as fh:
                data = json.load(fh)
        except (OSError, ValueError):
            return
        if not isinstance(data, dict):
            return

        saved_last_dir = data.get("last_output_dir")
        if isinstance(saved_last_dir, str) and saved_last_dir:
            self.last_output_dir = saved_last_dir
        try:
            self.dark_var.set(bool(data.get("dark_mode", False)))
        except Exception:
            pass
        try:
            saved_mode = data.get("escape_mode")
            if isinstance(saved_mode, str) and saved_mode in ("gs", "1c", "ozon"):
                self.escape_mode_var.set(saved_mode)
        except Exception:
            pass
        try:
            self.merge_csv_var.set(bool(data.get("merge_csv", False)))
        except Exception:
            pass

        # выбор сканера (сохраняем между запусками)
        try:
            smode = data.get("scanner_mode")
            if smode in ("keyboard", "com"):
                self.scanner_mode_var.set(smode)
        except Exception:
            pass
        try:
            sport = data.get("scanner_port")
            if isinstance(sport, str) and sport:
                self.scanner_port_var.set(sport)
        except Exception:
            pass

        # настройки вкладки CSV -> DataMatrix
        try:
            if isinstance(data.get("dm_label_size"), str):
                self.dm_label_size_var.set(data["dm_label_size"])
        except Exception:
            pass
        try:
            if isinstance(data.get("dm_group"), str):
                self.dm_group_var.set(data["dm_group"])
        except Exception:
            pass
        for key, attr in (
            ("dm_show_date", "dm_show_date_var"),
            ("dm_show_expiry", "dm_show_expiry_var"),
            ("dm_show_batch", "dm_show_batch_var"),
            ("dm_show_code", "dm_show_code_var"),
            ("dm_show_short", "dm_show_short_var"),
        ):
            try:
                if key in data:
                    getattr(self, attr).set(bool(data[key]))
            except Exception:
                pass

        # жирность и размеры шрифтов
        for plain_key, var_name in (
            ("dm_bold_name", "dm_bold_name_var"),
            ("dm_bold_date", "dm_bold_date_var"),
            ("dm_bold_expiry", "dm_bold_expiry_var"),
            ("dm_bold_batch", "dm_bold_batch_var"),
            ("dm_bold_code", "dm_bold_code_var"),
        ):
            try:
                if plain_key in data:
                    getattr(self, var_name).set(bool(data[plain_key]))
            except Exception:
                pass
        for plain_key, var_name in (
            ("dm_size_name", "dm_size_name_var"),
            ("dm_size_date", "dm_size_date_var"),
            ("dm_size_expiry", "dm_size_expiry_var"),
            ("dm_size_batch", "dm_size_batch_var"),
            ("dm_size_code", "dm_size_code_var"),
        ):
            try:
                if isinstance(data.get(plain_key), str):
                    getattr(self, var_name).set(data[plain_key])
            except Exception:
                pass

        # формат дат производства и срока годности
        for plain_key, var_name in (
            ("dm_date_format", "dm_date_format_var"),
            ("dm_expiry_format", "dm_expiry_format_var"),
        ):
            try:
                val = data.get(plain_key)
                if val in self.DATE_FORMATS:
                    getattr(self, var_name).set(val)
            except Exception:
                pass

        # сохранённые шаблоны этикеток
        saved_profiles = data.get("profiles")
        if isinstance(saved_profiles, dict):
            self._profiles = {
                str(name): dict(profile)
                for name, profile in saved_profiles.items()
                if isinstance(profile, dict)
            }

    def _save_config(self) -> None:
        cfg = config_path()
        data = {
            "last_output_dir": self.last_output_dir or "",
            "dark_mode": bool(self.dark_var.get()),
            "escape_mode": self.escape_mode_var.get(),
            "merge_csv": bool(self.merge_csv_var.get()),
            "dm_label_size": self.dm_label_size_var.get(),
            "dm_group": self.dm_group_var.get(),
            "dm_show_date": bool(self.dm_show_date_var.get()),
            "dm_show_expiry": bool(self.dm_show_expiry_var.get()),
            "dm_show_batch": bool(self.dm_show_batch_var.get()),
            "dm_show_code": bool(self.dm_show_code_var.get()),
            "dm_show_short": bool(self.dm_show_short_var.get()),
            "dm_bold_name": bool(self.dm_bold_name_var.get()),
            "dm_bold_date": bool(self.dm_bold_date_var.get()),
            "dm_bold_expiry": bool(self.dm_bold_expiry_var.get()),
            "dm_bold_batch": bool(self.dm_bold_batch_var.get()),
            "dm_bold_code": bool(self.dm_bold_code_var.get()),
            "dm_size_name": self.dm_size_name_var.get(),
            "dm_size_date": self.dm_size_date_var.get(),
            "dm_size_expiry": self.dm_size_expiry_var.get(),
            "dm_size_batch": self.dm_size_batch_var.get(),
            "dm_size_code": self.dm_size_code_var.get(),
            "dm_date_format": self.dm_date_format_var.get(),
            "dm_expiry_format": self.dm_expiry_format_var.get(),
            "profiles": self._profiles,
            "scanner_mode": self.scanner_mode_var.get(),
            "scanner_port": self.scanner_port_var.get(),
        }
        try:
            with open(cfg, "w", encoding="utf-8") as fh:
                json.dump(data, fh, ensure_ascii=False, indent=2)
        except OSError:
            pass

    def _on_close(self) -> None:
        if self.running and not messagebox.askyesno(
            "Выход", "Конвертация выполняется. Остановить её и закрыть программу?"
        ):
            return
        self._save_config()
        self._stop_com_listener()
        self.cancel_event.set()
        self.root.destroy()


def _global_excepthook(exc_type, exc_value, exc_tb) -> None:
    """Глобальный перехват необработанных исключений: пишет в лог-файл и
    показывает окно с ошибкой и кнопкой «Скопировать»."""
    import traceback as _tb
    import tkinter as _tk
    text = "".join(_tb.format_exception(exc_type, exc_value, exc_tb))
    try:
        with open(log_file_path(), "a", encoding="utf-8") as _fh:
            _fh.write("\n===== НЕОБРАБОТАННАЯ ОШИБКА =====\n" + text + "\n")
    except Exception:
        pass
    try:
        root = getattr(_tk._default_root, "winfo_toplevel", lambda: None)()
    except Exception:
        root = None
    win_tk = root if root is not None else _tk.Tk()
    win_tk.withdraw()
    top = _tk.Toplevel(win_tk)
    top.title("Ошибка программы")
    top.transient(win_tk)
    top.grab_set()
    _tk.Label(top, text="Ошибка программы. Подробности ниже — скопируйте их "
                        "и отправьте разработчику:").pack(padx=12, pady=(10, 4))
    txt = _tk.Text(top, width=80, height=12, font=("Consolas", 9))
    txt.insert("1.0", text)
    txt.configure(state="disabled")
    txt.pack(padx=12, pady=4)
    btns = _tk.Frame(top)
    btns.pack(pady=8)
    def _copy() -> None:
        try:
            win_tk.clipboard_clear()
            win_tk.clipboard_append(text)
            _lbl.config(text="Скопировано в буфер обмена.")
        except Exception:
            pass
    _btn = _tk.Button(btns, text="Скопировать ошибку", command=_copy)
    _btn.pack(side="left", padx=6)
    _lbl = _tk.Label(btns, text="")
    _lbl.pack(side="left", padx=6)
    _tk.Button(btns, text="Закрыть", command=top.destroy).pack(side="left", padx=6)
    win_tk.deiconify()

def run_gui() -> None:
    import sys as _sys
    _sys.excepthook = _global_excepthook
    try:
        from tkinterdnd2 import TkinterDnD

        root = TkinterDnD.Tk()
    except Exception:
        root = tk.Tk()
    app = ConverterGui(root)
    root.mainloop()
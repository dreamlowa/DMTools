import sys


def early_update_check() -> None:
    """Ранняя проверка обновлений ДО построения основного GUI, в отдельном
    потоке. Сделано специально, чтобы обновление срабатывало даже если
    графический интерфейс сломан или падает при запуске: поток переживает
    краш GUI-потока и может предложить обновление в простом диалоге,
    не зависящем от основного окна."""
    try:
        import threading
        threading.Thread(target=_early_work, daemon=True).start()
    except Exception:
        pass


def _early_work() -> None:
    try:
        from updater import check_for_update, run_self_update
        import tkinter as tk
        from tkinter import messagebox
        import time

        # даём основному GUI время стартовать; если он падает — поток всё
        # равно продолжает и показывает обновление
        time.sleep(1.0)

        result = check_for_update()
        if result.get("status") != "ok":
            return
        update = result.get("update")
        if update is None:
            return

        # простой tk-диалог, не зависящий от экземпляра GUI
        root = tk.Tk()
        root.withdraw()
        notes = (update.get("notes") or "").strip()
        msg = (f"Доступна новая версия v{update.get('version')}.\n"
               f"Хотите скачать и установить её сейчас?")
        if notes:
            msg += f"\n\nЧто нового:\n{notes}"
        if not messagebox.askyesno("Обновление DMTools", msg):
            root.destroy()
            return
        root.destroy()
        err = run_self_update(update["url"])
        if err:
            root = tk.Tk()
            root.withdraw()
            messagebox.showerror("Обновление", err)
            root.destroy()
    except Exception:
        pass


if __name__ == "__main__":
    early_update_check()
    from gui import run_gui
    run_gui()

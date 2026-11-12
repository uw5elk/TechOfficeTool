#!/usr/bin/env python3
"""TechOfficeTool - графічний інтерфейс. Уся логіка живе в techoffice_core.py."""
import queue
import sys
import threading
from pathlib import Path

try:
    import tkinter as tk
    from tkinter import ttk, filedialog, messagebox
except ImportError:
    sys.exit(1)

try:
    import techoffice_core as core
except ImportError as _e:  # напр. не встановлено Pillow - покажемо зрозуміле повідомлення
    _root = tk.Tk()
    _root.withdraw()
    messagebox.showerror(
        "TechOfficeTool",
        "Не вдалося запустити програму:\n%s\n\nВстановіть залежності:\n"
        "pip install -r requirements.txt" % _e,
    )
    sys.exit(1)

from techoffice_core import (
    LEVELS,
    MODES,
    MODE_CONVERT,
    MODE_DOCX,
    MODE_IMAGES,
    MODE_RENAME,
    MODE_EXIF,
    SORT_OPTIONS,
    WIN32_OK,
    JobConfig,
)


class App(tk.Tk):
    def __init__(self):
        super().__init__()
        core.setup_logging()
        self.title("TechOfficeTool | UW5ELK")
        self.minsize(680, 400)
        self.configure(bg="#2d2d2d")
        style = ttk.Style(self)
        style.theme_use("clam")
        style.configure(".", background="#2d2d2d", foreground="white")
        style.configure("TButton", background="#555555", foreground="white")
        style.configure("TLabel", background="#2d2d2d", foreground="white")
        style.configure("TCheckbutton", background="#2d2d2d", foreground="white")
        style.configure("TProgressbar", background="#4a90e2", troughcolor="#3c3c3c", bordercolor="#3c3c3c")
        style.configure("TEntry", fieldbackground="white", foreground="black")
        style.configure("TCombobox", fieldbackground="white", foreground="black", background="white", arrowcolor="black")
        style.map("TCombobox", fieldbackground=[("readonly", "white")], foreground=[("readonly", "black")], selectbackground=[("readonly", "white")], selectforeground=[("readonly", "black")])
        style.configure("TSpinbox", fieldbackground="white", foreground="black", arrowcolor="black", background="white")

        self.mode_var = tk.StringVar(value=MODES[0])  # першим іде конвертація .doc -> .docx
        self.input_var = tk.StringVar()
        self.output_var = tk.StringVar()
        self.level_var = tk.StringVar(value=list(LEVELS.keys())[1])
        self.same_folder = tk.BooleanVar(value=False)
        self.running = False
        self.recursive_var = tk.BooleanVar(value=False)
        self.keep_exif_var = tk.BooleanVar(value=True)
        self.keep_icc_var = tk.BooleanVar(value=True)
        self.paranoid_var = tk.BooleanVar(value=False)
        self.backup_var = tk.BooleanVar(value=True)
        self.dry_run_var = tk.BooleanVar(value=True)
        self.delete_doc_var = tk.BooleanVar(value=False)
        self.template_var = tk.StringVar(value="IMG_{n:04d}")
        self.start_num_var = tk.IntVar(value=1)
        self.sort_var = tk.StringVar(value=SORT_OPTIONS[0])

        # обмін з робочим потоком: тільки через чергу та Event
        self._queue = queue.Queue()
        self._cancel = threading.Event()
        self._thread = None
        self._closing = False

        self._build_ui()
        self._on_mode_change()
        self.protocol("WM_DELETE_WINDOW", self._on_close)
        self.after(100, self._poll_queue)

    # ------------------------------------------------------------------ UI --
    def _build_ui(self):
        pad = {"padx": 12, "pady": 6}
        frm = ttk.Frame(self, padding=20)
        frm.pack(fill=tk.BOTH, expand=True)
        ttk.Label(frm, text="Режим роботи:", font=("", 10, "bold")).grid(row=0, column=0, columnspan=3, sticky="w", **pad)
        self.mode_cb = ttk.Combobox(frm, textvariable=self.mode_var, values=list(MODES), state="readonly", width=52)
        self.mode_cb.grid(row=1, column=0, columnspan=3, sticky="ew", **pad)
        self.mode_cb.bind("<<ComboboxSelected>>", lambda e: self._on_mode_change())
        self.lbl_in = ttk.Label(frm, text="Вхідна папка:", font=("", 10, "bold"))
        self.lbl_in.grid(row=2, column=0, columnspan=3, sticky="w", **pad)
        ttk.Entry(frm, textvariable=self.input_var, width=55).grid(row=3, column=0, columnspan=2, sticky="ew", **pad)
        ttk.Button(frm, text="Обрати…", command=self.pick_input).grid(row=3, column=2, **pad)
        self.lbl_out = ttk.Label(frm, text="Куди зберігати:", font=("", 10, "bold"))
        self.lbl_out.grid(row=4, column=0, columnspan=3, sticky="w", **pad)
        self.out_entry = ttk.Entry(frm, textvariable=self.output_var, width=55)
        self.out_entry.grid(row=5, column=0, columnspan=2, sticky="ew", **pad)
        self.out_btn = ttk.Button(frm, text="Обрати…", command=self.pick_output)
        self.out_btn.grid(row=5, column=2, **pad)
        self.chk_same = ttk.Checkbutton(frm, text="Зберігати в ту саму папку", variable=self.same_folder, command=self._toggle_output)
        self.chk_same.grid(row=6, column=0, columnspan=3, sticky="w", **pad)
        self.lbl_level = ttk.Label(frm, text="Ступінь стиснення:", font=("", 10, "bold"))
        self.lbl_level.grid(row=7, column=0, columnspan=3, sticky="w", **pad)
        self.level_cb = ttk.Combobox(frm, textvariable=self.level_var, values=list(LEVELS.keys()), state="readonly", width=52)
        self.level_cb.grid(row=8, column=0, columnspan=3, sticky="ew", **pad)

        self.img_opts = ttk.Frame(frm)
        self.img_opts.grid(row=9, column=0, columnspan=3, sticky="w", **pad)
        ttk.Checkbutton(self.img_opts, text="Включати підпапки", variable=self.recursive_var).pack(anchor="w")
        ttk.Checkbutton(self.img_opts, text="Зберігати EXIF", variable=self.keep_exif_var).pack(anchor="w")

        self.exif_opts = ttk.Frame(frm)
        self.exif_opts.grid(row=9, column=0, columnspan=3, sticky="w", **pad)
        ttk.Checkbutton(self.exif_opts, text="Включати підпапки", variable=self.recursive_var).pack(anchor="w")
        ttk.Checkbutton(self.exif_opts, text="Зберігати кольори", variable=self.keep_icc_var).pack(anchor="w")
        ttk.Checkbutton(self.exif_opts, text="Повна переупаковка (максимальна очистка)", variable=self.paranoid_var).pack(anchor="w")
        ttk.Checkbutton(self.exif_opts, text="Резервна копія .bak (там EXIF/GPS лишається!)", variable=self.backup_var).pack(anchor="w")

        self.ren_opts = ttk.Frame(frm)
        self.ren_opts.grid(row=9, column=0, columnspan=3, sticky="w", **pad)
        ttk.Checkbutton(self.ren_opts, text="Включати підпапки", variable=self.recursive_var).pack(anchor="w")
        row = ttk.Frame(self.ren_opts); row.pack(anchor="w", pady=4)
        ttk.Label(row, text="Нова назва:").pack(side=tk.LEFT)
        ttk.Entry(row, textvariable=self.template_var, width=30).pack(side=tk.LEFT, padx=6)
        row2 = ttk.Frame(self.ren_opts); row2.pack(anchor="w", pady=4)
        ttk.Label(row2, text="Старт:").pack(side=tk.LEFT)
        ttk.Spinbox(row2, from_=0, to=999999, width=8, textvariable=self.start_num_var).pack(side=tk.LEFT, padx=6)
        row3 = ttk.Frame(self.ren_opts); row3.pack(anchor="w", pady=4)
        ttk.Label(row3, text="Сортування:").pack(side=tk.LEFT)
        ttk.Combobox(row3, textvariable=self.sort_var, values=list(SORT_OPTIONS), state="readonly", width=25).pack(side=tk.LEFT, padx=6)
        ttk.Checkbutton(self.ren_opts, text="Тільки попередній перегляд", variable=self.dry_run_var).pack(anchor="w", pady=4)
        ttk.Button(self.ren_opts, text="Відмінити перейменування…", command=self.undo_rename).pack(anchor="w", pady=4)

        self.conv_opts = ttk.Frame(frm)
        self.conv_opts.grid(row=9, column=0, columnspan=3, sticky="w", **pad)
        ttk.Checkbutton(self.conv_opts, text="Включати підпапки", variable=self.recursive_var).pack(anchor="w")
        ttk.Checkbutton(self.conv_opts, text="Видаляти оригінальні .doc після конвертації (у кошик)", variable=self.delete_doc_var).pack(anchor="w")

        self.lbl_convert_hint = ttk.Label(frm, text="Потрібен Microsoft Word для конвертації.", foreground="#aaaaaa")
        self.lbl_convert_hint.grid(row=10, column=0, columnspan=3, sticky="w", **pad)
        self.status = ttk.Label(frm, text="Оберіть режим і натисніть «Старт»", foreground="#aaaaaa")
        self.status.grid(row=11, column=0, columnspan=3, sticky="w", **pad)
        self.progress = ttk.Progressbar(frm, mode="determinate", length=400)
        self.progress.grid(row=12, column=0, columnspan=3, sticky="ew", **pad)
        self.detail = ttk.Label(frm, text="", foreground="#aaaaaa")
        self.detail.grid(row=13, column=0, columnspan=3, sticky="w", **pad)
        btn_frm = ttk.Frame(frm)
        btn_frm.grid(row=14, column=0, columnspan=3, pady=16)
        self.start_btn = ttk.Button(btn_frm, text="▶ Старт", command=self.start, width=16)
        self.start_btn.pack(side=tk.LEFT, padx=8)
        self.cancel_btn = ttk.Button(btn_frm, text="■ Скасувати", command=self.cancel, width=14, state="disabled")
        self.cancel_btn.pack(side=tk.LEFT, padx=8)
        ttk.Button(btn_frm, text="Вихід", command=self._on_close, width=12).pack(side=tk.LEFT, padx=8)
        frm.columnconfigure(0, weight=1)

    def _hide_all_blocks(self):
        for w in (self.lbl_level, self.level_cb, self.img_opts, self.exif_opts, self.ren_opts, self.conv_opts, self.lbl_convert_hint):
            w.grid_remove()

    def _on_mode_change(self):
        self._hide_all_blocks()
        mode = self.mode_var.get()
        if mode == MODE_CONVERT:
            self.conv_opts.grid(); self.lbl_convert_hint.grid(); self.lbl_in.configure(text="Папка з .doc:"); self._show_output(True)
        elif mode == MODE_DOCX:
            self.lbl_level.grid(); self.level_cb.grid(); self.lbl_in.configure(text="Папка .docx (разом з підпапками):"); self._show_output(True)
        elif mode == MODE_IMAGES:
            self.lbl_level.grid(); self.level_cb.grid(); self.img_opts.grid(); self.lbl_in.configure(text="Папка зображень:"); self._show_output(True)
        elif mode == MODE_RENAME:
            self.ren_opts.grid(); self.lbl_in.configure(text="Папка фото:"); self._show_output(False)
        elif mode == MODE_EXIF:
            self.exif_opts.grid(); self.lbl_in.configure(text="Папка зображень:"); self._show_output(False)
        self.update_idletasks()
        self.geometry("")
        self._toggle_output()

    def _show_output(self, show):
        if show:
            self.lbl_out.grid(); self.out_entry.grid(); self.out_btn.grid(); self.chk_same.grid()
        else:
            self.lbl_out.grid_remove(); self.out_entry.grid_remove(); self.out_btn.grid_remove(); self.chk_same.grid_remove()

    def _toggle_output(self):
        if self.same_folder.get():
            if self.input_var.get():
                self.output_var.set(self.input_var.get())
            self.out_entry.configure(state="disabled"); self.out_btn.configure(state="disabled")
        else:
            self.out_entry.configure(state="normal"); self.out_btn.configure(state="normal")

    def pick_input(self):
        path = filedialog.askdirectory(title="Вхідна папка")
        if path:
            self.input_var.set(path)
            if self.same_folder.get():
                self.output_var.set(path)

    def pick_output(self):
        path = filedialog.askdirectory(title="Вихідна папка")
        if path:
            self.output_var.set(path)

    # ------------------------------------------------------------- запуск --
    def start(self):
        if self.running:
            return
        inp = self.input_var.get().strip()
        if not inp or not Path(inp).is_dir():
            messagebox.showerror("Помилка", "Оберіть вхідну папку.")
            return
        mode = self.mode_var.get()
        if mode in (MODE_CONVERT, MODE_DOCX, MODE_IMAGES):
            if self.same_folder.get():
                out = inp
            else:
                out = self.output_var.get().strip()
                if not out:
                    messagebox.showerror("Помилка", "Оберіть вихідну папку.")
                    return
                Path(out).mkdir(parents=True, exist_ok=True)
        else:
            out = inp
        if mode == MODE_CONVERT and not WIN32_OK:
            messagebox.showerror("Помилка", "Для конвертації потрібен pywin32.")
            return
        try:
            start_num = int(self.start_num_var.get())
        except Exception:
            messagebox.showerror("Помилка", "Поле «Старт» має бути цілим числом.")
            return
        if mode == MODE_RENAME:
            bad = core.validate_template(self.template_var.get().strip() or "IMG_{n:04d}")
            if bad:
                messagebox.showerror("Помилка", bad)
                return
        # попередження перед незворотними діями
        if mode in (MODE_DOCX, MODE_IMAGES) and out == inp:
            if not messagebox.askyesno("Підтвердження", "Файли буде замінено на місці (оригінали не збережуться).\nПродовжити?"):
                return
        if mode == MODE_EXIF and not self.backup_var.get():
            if not messagebox.askyesno("Підтвердження", "Резервні копії вимкнено - EXIF буде видалено безповоротно.\nПродовжити?"):
                return

        cfg = JobConfig(
            mode=mode,
            input_dir=Path(inp),
            output_dir=Path(out),
            level=LEVELS[self.level_var.get()] if mode in (MODE_DOCX, MODE_IMAGES) else None,
            recursive=self.recursive_var.get(),
            keep_exif=self.keep_exif_var.get(),
            keep_icc=self.keep_icc_var.get(),
            paranoid=self.paranoid_var.get(),
            backup=self.backup_var.get(),
            dry_run=self.dry_run_var.get(),
            delete_doc=self.delete_doc_var.get(),
            template=self.template_var.get(),
            start_num=start_num,
            sort=self.sort_var.get(),
        )
        self.running = True
        self._cancel.clear()
        self.start_btn.configure(state="disabled")
        self.cancel_btn.configure(state="normal")
        self.status.configure(text="Пошук...")
        self.progress["value"] = 0
        self.detail.configure(text="")
        self._thread = threading.Thread(target=self._worker, args=(cfg,), daemon=True)
        self._thread.start()

    def cancel(self):
        if self.running:
            self._cancel.set()
            self.cancel_btn.configure(state="disabled")
            self.status.configure(text="Скасування після поточного файлу…")

    def _worker(self, cfg):
        """Працює у потоці й не торкається віджетів - лише кладе повідомлення в чергу."""
        try:
            msg = core.run_job(
                cfg,
                progress=lambda i, t, n: self._queue.put(("progress", i, t, n)),
                cancel=self._cancel,
            )
        except Exception as e:
            core.LOG.exception("Збій завдання")
            msg = "Сталася помилка: %s: %s" % (type(e).__name__, e)
            if core.LOG_PATH:
                msg += "\nДеталі у журналі:\n%s" % core.LOG_PATH
        self._queue.put(("done", msg))

    def _poll_queue(self):
        try:
            while True:
                item = self._queue.get_nowait()
                if item[0] == "progress":
                    self._upd(*item[1:])
                elif item[0] == "done":
                    self._done(item[1])
        except queue.Empty:
            pass
        if not self._closing:
            self.after(100, self._poll_queue)

    def _upd(self, i, total, name):
        self.progress.configure(maximum=total)
        self.progress["value"] = i
        self.status.configure(text="Обробка %d з %d" % (i, total))
        self.detail.configure(text=name)

    def _done(self, msg):
        self.running = False
        self.start_btn.configure(state="normal")
        self.cancel_btn.configure(state="disabled")
        self.status.configure(text="Завершено")
        self.detail.configure(text="")
        if core.LOG_PATH and ("Помил" in msg or "помил" in msg):
            msg += "\n\nЖурнал: %s" % core.LOG_PATH
        messagebox.showinfo("Результат", msg)

    # ------------------------------------------------------ відміна rename --
    def undo_rename(self):
        path = filedialog.askopenfilename(
            title="Журнал перейменування",
            initialdir=self.input_var.get() or None,
            filetypes=[("Журнал перейменування", "*.json"), ("Усі файли", "*.*")],
        )
        if not path:
            return
        try:
            restored, skipped, errors = core.undo_rename(Path(path))
        except Exception as e:
            core.LOG.exception("Не вдалося відмінити перейменування")
            messagebox.showerror("Помилка", "Не вдалося прочитати журнал:\n%s" % e)
            return
        msg = "Повернуто назв: %d\nПропущено: %d" % (restored, skipped)
        if errors:
            msg += "\nПомилок: %d\n\n" % len(errors) + "\n".join(errors[:3])
        messagebox.showinfo("Результат", msg)

    # ------------------------------------------------------------- вихід --
    def _on_close(self):
        if self.running:
            if not messagebox.askyesno("Вихід", "Триває обробка. Скасувати її та вийти?"):
                return
            self._cancel.set()
            self._closing = True
            self.status.configure(text="Завершую поточний файл…")
            self.after(200, self._wait_and_close)
        else:
            self._closing = True
            self.destroy()

    def _wait_and_close(self):
        if self._thread is not None and self._thread.is_alive():
            self.after(200, self._wait_and_close)
        else:
            self.destroy()


if __name__ == "__main__":
    app = App()
    app.mainloop()

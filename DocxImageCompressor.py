#!/usr/bin/env python3
import os, sys, json, shutil, zipfile, tempfile, threading, traceback, time
from datetime import datetime
from pathlib import Path
try:
    import tkinter as tk
    from tkinter import ttk, filedialog, messagebox
except ImportError:
    sys.exit(1)
try:
    from PIL import Image, ImageFile
    ImageFile.LOAD_TRUNCATED_IMAGES = True
except ImportError:
    sys.exit(1)
try:
    import pythoncom, win32com.client
    WIN32_OK = True
except ImportError:
    WIN32_OK = False

LEVELS = {
    "Максимальна якість (архів)": {"max_side": 4000, "quality": 95, "png2jpg": False},
    "Для друку (≈220 ppi)": {"max_side": 2400, "quality": 88, "png2jpg": True},
    "Для екрану (≈150 ppi)": {"max_side": 1600, "quality": 82, "png2jpg": True},
    "Сильне стиснення": {"max_side": 1200, "quality": 75, "png2jpg": True},
}
MODES = ("Стиснути зображення в .docx", "Конвертувати .doc → .docx", "Стиснути звичайні зображення", "Перейменувати фото", "Видалити EXIF")
IMG_EXTS = {".jpg", ".jpeg", ".png", ".gif", ".bmp", ".tiff", ".tif", ".webp"}
SORT_OPTIONS = ("За іменем", "За датою EXIF", "За датою файлу", "Без сортування")

def is_valid_docx(path: Path) -> bool:
    try:
        with open(path, "rb") as f: return f.read(4) == b"PK\x03\x04"
    except Exception: return False

def _get_exif_datetime(im: Image.Image):
    try:
        exif = im.getexif()
        dt = exif.get(306)
        if dt: return datetime.strptime(dt, "%Y:%m:%d %H:%M:%S")
    except Exception: pass
    return None

def compress_image_file(img_path: Path, max_side: int, quality: int, png2jpg: bool, keep_exif: bool = True) -> bool:
    try:
        with Image.open(img_path) as im:
            exif_bytes = im.info.get("exif") if keep_exif else None
            has_transparency = im.mode in ("RGBA", "LA") or (im.mode == "P" and "transparency" in im.info)
            if not has_transparency and im.mode != "RGB": im = im.convert("RGB")
            w, h = im.size
            if max(w, h) > max_side:
                ratio = max_side / max(w, h)
                im = im.resize((int(w * ratio), int(h * ratio)), Image.Resampling.LANCZOS)
            if has_transparency or (img_path.suffix.lower() == ".png" and not png2jpg):
                im.save(img_path, "PNG", optimize=True, compress_level=6)
            else:
                kwargs = {"quality": quality, "optimize": True, "progressive": True}
                if exif_bytes: kwargs["exif"] = exif_bytes
                im.save(img_path, "JPEG", **kwargs)
        return True
    except Exception: return False

def process_one_image(src: Path, dst: Path, max_side: int, quality: int, png2jpg: bool, keep_exif: bool) -> dict:
    res = {"ok": False, "old": 0, "new": 0}
    try:
        res["old"] = src.stat().st_size
        dst.parent.mkdir(parents=True, exist_ok=True)
        if str(src) != str(dst): shutil.copy2(src, dst)
        compress_image_file(dst, max_side, quality, png2jpg, keep_exif)
        res["new"] = dst.stat().st_size
        res["ok"] = True
    except Exception: pass
    return res

def _get_date_for_sort(p: Path, mode: str):
    if mode.startswith("За датою EXIF"):
        try:
            with Image.open(p) as im:
                dt = _get_exif_datetime(im)
                if dt: return dt.timestamp()
        except Exception: pass
        return p.stat().st_mtime
    if mode.startswith("За датою файлу"): return p.stat().st_mtime
    return 0

def _get_exif_date_str(p: Path) -> str:
    try:
        with Image.open(p) as im:
            dt = _get_exif_datetime(im)
            if dt: return dt.strftime("%Y-%m-%d")
    except Exception: pass
    return datetime.fromtimestamp(p.stat().st_mtime).strftime("%Y-%m-%d")

def _format_name(template: str, n: int, src: Path) -> str:
    try:
        return template.format(n=n, name=src.stem, parent=src.parent.name, date=_get_exif_date_str(src), ext=src.suffix.lstrip("."))
    except Exception: return f"{src.stem}_{n}"

def strip_exif_one(p: Path, keep_icc: bool, paranoid: bool, backup: bool) -> str:
    try:
        with Image.open(p) as im:
            had_meta = bool(im.getexif()) or any(k in im.info for k in ("exif", "icc_profile", "comment", "xmp"))
            if not had_meta: return "already_clean"
            if backup:
                bak = p.with_suffix(p.suffix + ".bak")
                if not bak.exists(): shutil.copy2(p, bak)
            icc = im.info.get("icc_profile") if keep_icc else None
            fmt = im.format or "JPEG"
            has_alpha = im.mode in ("RGBA", "LA") or (im.mode == "P" and "transparency" in im.info)
            clean = Image.new("RGBA" if has_alpha else "RGB", im.size)
            if paranoid: clean.paste(im.convert(clean.mode))
            else: clean = im.copy()
            save_kwargs = {}
            if fmt in ("JPEG", "JPG"): save_kwargs = {"quality": 95, "optimize": True}
            elif fmt == "PNG": save_kwargs = {"optimize": True, "compress_level": 9}
            elif fmt == "WEBP": save_kwargs = {"quality": 95, "method": 6}
            if icc and fmt in ("JPEG", "PNG", "WEBP"): save_kwargs["icc_profile"] = icc
            clean.save(p, fmt, **save_kwargs)
        return "cleaned"
    except Exception: return "error"

def compress_image(img_path: Path, max_side: int, quality: int, png2jpg: bool) -> bool:
    return compress_image_file(img_path, max_side, quality, png2jpg, keep_exif=False)

def process_one_docx(src: Path, dst: Path, max_side: int, quality: int, png2jpg: bool) -> dict:
    result = {"ok": False, "images": 0, "old": 0, "new": 0}
    temp_out = None
    try:
        result["old"] = src.stat().st_size
        same = str(src) == str(dst)
        write_path = temp_out = dst.with_suffix(".tmp_c.docx") if same else dst
        with tempfile.TemporaryDirectory() as tmp:
            tmp_path = Path(tmp)
            with zipfile.ZipFile(src, "r") as zf: zf.extractall(tmp_path)
            media = tmp_path / "word" / "media"
            if media.exists():
                for img in media.iterdir():
                    if img.suffix.lower() in IMG_EXTS and img.is_file():
                        if compress_image(img, max_side, quality, png2jpg): result["images"] += 1
            write_path.parent.mkdir(parents=True, exist_ok=True)
            with zipfile.ZipFile(write_path, "w", compression=zipfile.ZIP_DEFLATED) as zf:
                for root, _, files in os.walk(tmp_path):
                    for f in files:
                        full = Path(root) / f
                        zf.write(full, full.relative_to(tmp_path))
        if same and temp_out and temp_out.exists(): temp_out.replace(dst)
        result["new"] = dst.stat().st_size
        result["ok"] = True
    except Exception:
        if temp_out and temp_out.exists():
            try: temp_out.unlink()
            except Exception: pass
        if str(src) != str(dst):
            try:
                shutil.copy2(src, dst)
                result["new"] = dst.stat().st_size
            except Exception: pass
    return result

def convert_doc_to_docx(src: Path, dst: Path) -> dict:
    result = {"ok": False, "action": "fail", "old": 0, "new": 0}
    word, doc = None, None
    try:
        result["old"] = src.stat().st_size
        pythoncom.CoInitialize()
        word = win32com.client.DispatchEx("Word.Application")
        word.Visible = False
        word.DisplayAlerts = 0
        doc = word.Documents.Open(FileName=str(src), ConfirmConversions=False, ReadOnly=False, AddToRecentFiles=False, NoEncodingDialog=True)
        out_file = dst.with_suffix(".docx")
        out_file.parent.mkdir(parents=True, exist_ok=True)
        doc.SaveAs(FileName=str(out_file.resolve()), FileFormat=16)
        doc.Close(SaveChanges=False)
        doc = None
        result["new"] = out_file.stat().st_size if out_file.exists() else 0
        result["ok"] = True
        result["action"] = "convert"
        return result
    except Exception:
        try:
            out_file = dst.parent / src.name
            out_file.parent.mkdir(parents=True, exist_ok=True)
            shutil.copy2(src, out_file)
            result["new"] = out_file.stat().st_size
        except Exception: pass
        return result
    finally:
        if doc:
            try: doc.Close(SaveChanges=False)
            except Exception: pass
        if word:
            try: word.Quit()
            except Exception: pass
        try: pythoncom.CoUninitialize()
        except Exception: pass

def find_files_for_convert(root: Path) -> list:
    return [f for f in root.rglob("*") if f.is_file() and not f.name.startswith("~$") and (f.suffix.lower() == ".doc" or (f.suffix.lower() == ".docx" and not is_valid_docx(f)))]

class App(tk.Tk):
    def __init__(self):
        super().__init__()
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
        
        self.mode_var = tk.StringVar(value=MODES[0])
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
        self.template_var = tk.StringVar(value="IMG_{n:04d}")
        self.start_num_var = tk.IntVar(value=1)
        self.step_var = tk.IntVar(value=1)
        self.sort_var = tk.StringVar(value=SORT_OPTIONS[0])
        self._build_ui()
        self._on_mode_change()

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
        ttk.Checkbutton(self.exif_opts, text="Створювати резервну копію", variable=self.backup_var).pack(anchor="w")
        
        self.ren_opts = ttk.Frame(frm)
        self.ren_opts.grid(row=9, column=0, columnspan=3, sticky="w", **pad)
        ttk.Checkbutton(self.ren_opts, text="Включати підпапки", variable=self.recursive_var).pack(anchor="w")
        row = ttk.Frame(self.ren_opts); row.pack(anchor="w", pady=4)
        ttk.Label(row, text="Нова назва:").pack(side=tk.LEFT)
        ttk.Entry(row, textvariable=self.template_var, width=30).pack(side=tk.LEFT, padx=6)
        row2 = ttk.Frame(self.ren_opts); row2.pack(anchor="w", pady=4)
        ttk.Label(row2, text="Старт:").pack(side=tk.LEFT)
        ttk.Spinbox(row2, from_=0, to=999999, width=8, textvariable=self.start_num_var).pack(side=tk.LEFT, padx=6)
        ttk.Label(row2, text="Крок:").pack(side=tk.LEFT, padx=6)
        ttk.Spinbox(row2, from_=1, to=10000, width=6, textvariable=self.step_var).pack(side=tk.LEFT, padx=6)
        row3 = ttk.Frame(self.ren_opts); row3.pack(anchor="w", pady=4)
        ttk.Label(row3, text="Сортування:").pack(side=tk.LEFT)
        ttk.Combobox(row3, textvariable=self.sort_var, values=list(SORT_OPTIONS), state="readonly", width=25).pack(side=tk.LEFT, padx=6)
        ttk.Checkbutton(self.ren_opts, text="Тільки попередній перегляд", variable=self.dry_run_var).pack(anchor="w", pady=4)
        
        self.lbl_convert_hint = ttk.Label(frm, text="Потрібен Microsoft Word для конвертації.", foreground="#aaaaaa")
        self.lbl_convert_hint.grid(row=9, column=0, columnspan=3, sticky="w", **pad)
        self.status = ttk.Label(frm, text="Оберіть режим і натисніть «Старт»", foreground="#333")
        self.status.grid(row=10, column=0, columnspan=3, sticky="w", **pad)
        self.progress = ttk.Progressbar(frm, mode="determinate", length=400)
        self.progress.grid(row=11, column=0, columnspan=3, sticky="ew", **pad)
        self.detail = ttk.Label(frm, text="", foreground="#aaaaaa")
        self.detail.grid(row=12, column=0, columnspan=3, sticky="w", **pad)
        btn_frm = ttk.Frame(frm)
        btn_frm.grid(row=13, column=0, columnspan=3, pady=16)
        self.start_btn = ttk.Button(btn_frm, text="▶ Старт", command=self.start, width=16)
        self.start_btn.pack(side=tk.LEFT, padx=8)
        ttk.Button(btn_frm, text="Вихід", command=self.destroy, width=12).pack(side=tk.LEFT, padx=8)
        frm.columnconfigure(0, weight=1)

    def _hide_all_blocks(self):
        for w in (self.lbl_level, self.level_cb, self.img_opts, self.exif_opts, self.ren_opts, self.lbl_convert_hint): w.grid_remove()
    
    def _on_mode_change(self):
        self._hide_all_blocks()
        mode = self.mode_var.get()
        if mode == MODES[0]:
            self.lbl_level.grid(); self.level_cb.grid(); self.lbl_in.configure(text="Папка .docx:"); self._show_output(True)
        elif mode == MODES[1]:
            self.lbl_convert_hint.grid(); self.lbl_in.configure(text="Папка .doc/.docx:"); self._show_output(True)
        elif mode == MODES[2]:
            self.lbl_level.grid(); self.level_cb.grid(); self.img_opts.grid(); self.lbl_in.configure(text="Папка зображень:"); self._show_output(True)
        elif mode == MODES[3]:
            self.ren_opts.grid(); self.lbl_in.configure(text="Папка фото:"); self._show_output(False)
        elif mode == MODES[4]:
            self.exif_opts.grid(); self.lbl_in.configure(text="Папка зображень:"); self._show_output(False)
        self.update_idletasks()
        self.geometry("")
        self._toggle_output()

    def _show_output(self, show):
        if show: self.lbl_out.grid(); self.out_entry.grid(); self.out_btn.grid(); self.chk_same.grid()
        else: self.lbl_out.grid_remove(); self.out_entry.grid_remove(); self.out_btn.grid_remove(); self.chk_same.grid_remove()

    def _toggle_output(self):
        if self.same_folder.get():
            if self.input_var.get(): self.output_var.set(self.input_var.get())
            self.out_entry.configure(state="disabled"); self.out_btn.configure(state="disabled")
        else:
            self.out_entry.configure(state="normal"); self.out_btn.configure(state="normal")

    def pick_input(self):
        path = filedialog.askdirectory(title="Вхідна папка")
        if path:
            self.input_var.set(path)
            if self.same_folder.get(): self.output_var.set(path)

    def pick_output(self):
        path = filedialog.askdirectory(title="Вихідна папка")
        if path: self.output_var.set(path)

    def start(self):
        if self.running: return
        inp = self.input_var.get().strip()
        if not inp or not Path(inp).is_dir():
            messagebox.showerror("Помилка", "Оберіть вхідну папку."); return
        mode = self.mode_var.get()
        if mode in (MODES[0], MODES[1], MODES[2]):
            if self.same_folder.get(): out = inp
            else:
                out = self.output_var.get().strip()
                if not out: messagebox.showerror("Помилка", "Оберіть вихідну папку."); return
                Path(out).mkdir(parents=True, exist_ok=True)
        else: out = inp
        if mode == MODES[1] and not WIN32_OK:
            messagebox.showerror("Помилка", "Для конвертації потрібен pywin32."); return
        level = LEVELS[self.level_var.get()] if mode in (MODES[0], MODES[2]) else None
        self.running = True; self.start_btn.configure(state="disabled"); self.status.configure(text="Пошук..."); self.progress["value"] = 0; self.detail.configure(text="")
        threading.Thread(target=self._worker, args=(mode, inp, out, level), daemon=True).start()

    def _worker(self, mode, inp, out, level):
        input_path, output_path = Path(inp), Path(out)
        if mode == MODES[0]:
            files = [f for f in input_path.rglob("*.docx") if f.is_file() and not f.name.startswith("~$") and not f.name.endswith(".tmp_c.docx") and is_valid_docx(f)]
            if not files: self.after(0, lambda: self._done("Не знайдено .docx")); return
            total = len(files); self.after(0, lambda: self.progress.configure(maximum=total))
            old_sum = new_sum = img_sum = errs = 0
            for i, src in enumerate(files, 1):
                dst = output_path / src.relative_to(input_path)
                self.after(0, lambda i=i, t=total, n=src.name[:50]: self._upd(i, t, n))
                res = process_one_docx(src, dst, level["max_side"], level["quality"], level["png2jpg"])
                old_sum += res["old"]; new_sum += res["new"]; img_sum += res["images"]
                if not res["ok"]: errs += 1
            msg = f"Готово! Файлів: {total}, Зображень: {img_sum}\nБуло: {old_sum/1024/1024:.1f} MB, Стало: {new_sum/1024/1024:.1f} MB"
            if old_sum > 0: msg += f"\nЗменшення: {(1 - new_sum / old_sum) * 100:.1f}%"
            if errs: msg += f"\nПомилок: {errs}"
            self.after(0, lambda: self._done(msg))
        elif mode == MODES[1]:
            files = find_files_for_convert(input_path)
            if not files: self.after(0, lambda: self._done("Немає .doc")); return
            total = len(files); self.after(0, lambda: self.progress.configure(maximum=total))
            ok, fail = 0, 0
            for i, src in enumerate(files, 1):
                dst = output_path / (src.relative_to(input_path) if src.is_relative_to(input_path) else src.name)
                self.after(0, lambda i=i, t=total, n=src.name[:50]: self._upd(i, t, n))
                res = convert_doc_to_docx(src, dst)
                if res["ok"]: ok += 1
                else: fail += 1
            self.after(0, lambda: self._done(f"Готово! Конвертовано: {ok}, Помилок: {fail}"))
        elif mode == MODES[2]:
            rec = self.recursive_var.get(); keep_exif = self.keep_exif_var.get()
            files = [f for f in (input_path.rglob("*") if rec else input_path.glob("*")) if f.is_file() and f.suffix.lower() in IMG_EXTS]
            if not files: self.after(0, lambda: self._done("Немає зображень")); return
            total = len(files); self.after(0, lambda: self.progress.configure(maximum=total))
            old_sum = new_sum = ok = err = 0
            for i, src in enumerate(files, 1):
                dst = output_path / src.relative_to(input_path)
                self.after(0, lambda i=i, t=total, n=src.name[:50]: self._upd(i, t, n))
                res = process_one_image(src, dst, level["max_side"], level["quality"], level["png2jpg"], keep_exif)
                if res["ok"]: ok += 1; old_sum += res["old"]; new_sum += res["new"]
                else: err += 1
            msg = f"Готово! Оброблено: {ok}/{total}\nБуло: {old_sum/1024/1024:.1f} MB, Стало: {new_sum/1024/1024:.1f} MB"
            if old_sum > 0: msg += f"\nЗменшення: {(1 - new_sum / old_sum) * 100:.1f}%"
            if err: msg += f"\nПомилок: {err}"
            self.after(0, lambda: self._done(msg))
        elif mode == MODES[3]:
            tmpl = self.template_var.get().strip() or "IMG_{n:04d}"
            dry = self.dry_run_var.get(); rec = self.recursive_var.get()
            self.after(0, lambda: self.status.configure(text="Аналіз..."))
            
            # Оновлена логіка перейменування (без зламаних функцій)
            files = [f for f in (input_path.rglob("*") if rec else input_path.glob("*")) if f.is_file() and f.suffix.lower() in IMG_EXTS]
            if not files:
                self.after(0, lambda: self._done("У вибраній папці не знайдено зображень.")); return
            
            if self.sort_var.get() == "За іменем": files.sort(key=lambda p: p.name.lower())
            elif self.sort_var.get().startswith("За датою"): files.sort(key=lambda p: _get_date_for_sort(p, self.sort_var.get()))
            
            used = {}
            n = self.start_num_var.get()
            planned = []
            for src in files:
                base = _format_name(tmpl, n, src)
                ext = src.suffix
                candidate = src.parent / f"{base}{ext}"
                k = 1
                while (candidate.exists() and str(candidate) != str(src)) or str(candidate).lower() in used:
                    candidate = src.parent / f"{base}_{k}{ext}"
                    k += 1
                used[str(candidate).lower()] = True
                planned.append((src, candidate))
                n += self.step_var.get()
                
            if dry:
                lines = [f"Знайдено: {len(planned)}", "План (перші 50):"]
                for s, d in planned[:50]: lines.append(f"{s.name} → {d.name}")
                if len(planned) > 50: lines.append(f"... ще {len(planned)-50}")
                self.after(0, lambda: self._done("\n".join(lines))); return

            total = len(planned); self.after(0, lambda: self.progress.configure(maximum=total))
            ok, err, mapping = 0, 0, []
            errors_text = []
            
            for i, (src, dst) in enumerate(planned, 1):
                self.after(0, lambda i=i, t=total, n=src.name[:50]: self._upd(i, t, n))
                try:
                    if str(src) != str(dst):
                        success = False
                        for attempt in range(5):
                            try:
                                os.rename(str(src), str(dst))
                                mapping.append({"old": str(src), "new": str(dst)})
                                success = True
                                break
                            except FileNotFoundError:
                                time.sleep(0.5) # Даємо Windows більше часу (0.5с)
                            except Exception:
                                time.sleep(0.5)
                        
                        if success:
                            ok += 1
                        else:
                            err += 1
                            errors_text.append(f"{src.name}: Файл заблоковано системою")
                    else:
                        ok += 1
                except Exception as e:
                    err += 1
                    errors_text.append(f"{src.name}: {e}")
            
            if mapping:
                try:
                    with open(input_path / f".rename_log_{datetime.now():%Y%m%d_%H%M%S}.json", "w", encoding="utf-8") as f: 
                        json.dump({"time": datetime.now().isoformat(), "mapping": mapping}, f, ensure_ascii=False, indent=2)
                except Exception: pass
            msg = f"Перейменовано: {ok}\nПомилок: {err}"
            if mapping: msg += "\nЛог збережено (.rename_log_*.json)"
            if errors_text: msg += "\n\nДеталі помилок:\n" + "\n".join(errors_text[:3])
            self.after(0, lambda: self._done(msg))
        elif mode == MODES[4]:
            rec = self.recursive_var.get(); keep_icc = self.keep_icc_var.get(); paranoid = self.paranoid_var.get(); backup = self.backup_var.get()
            files = [f for f in (input_path.rglob("*") if rec else input_path.glob("*")) if f.is_file() and f.suffix.lower() in IMG_EXTS]
            if not files: self.after(0, lambda: self._done("Немає зображень")); return
            total = len(files); self.after(0, lambda: self.progress.configure(maximum=total))
            cleaned = already = errs = 0
            for i, src in enumerate(files, 1):
                self.after(0, lambda i=i, t=total, n=src.name[:50]: self._upd(i, t, n))
                r = strip_exif_one(src, keep_icc, paranoid, backup)
                if r == "cleaned": cleaned += 1
                elif r == "already_clean": already += 1
                else: errs += 1
            msg = f"Очищено: {cleaned}\nБез EXIF: {already}\nПомилок: {errs}\nВсього: {total}"
            if backup: msg += "\nОригінали збережено як *.bak"
            self.after(0, lambda: self._done(msg))

    def _upd(self, i, total, name):
        self.progress["value"] = i; self.status.configure(text=f"Обробка {i} з {total}"); self.detail.configure(text=name)

    def _done(self, msg):
        self.running = False; self.start_btn.configure(state="normal"); self.status.configure(text="Завершено"); self.detail.configure(text="")
        messagebox.showinfo("Результат", msg)

if __name__ == "__main__":
    app = App()
    app.mainloop()

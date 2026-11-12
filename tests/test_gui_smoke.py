"""Димовий тест GUI без справжнього tkinter (у CI/Linux його часто немає).

Підставляємо мінімальну заглушку tkinter, імпортуємо DocxImageCompressor і
проганяємо побудову вікна, перемикання режимів, запуск завдання через чергу,
скасування та вихід. Це ловить NameError/AttributeError у GUI-коді, але не
перевіряє зовнішній вигляд - його треба глянути вручну у Windows.
"""
import importlib
import sys
import threading
import types
from pathlib import Path

import pytest
from PIL import Image


class Var:
    default = None

    def __init__(self, value=None, **kw):
        self._v = self.default if value is None else value

    def get(self):
        return self._v

    def set(self, v):
        self._v = v


class StringVar(Var):
    default = ""  # як у справжнього tkinter


class BooleanVar(Var):
    default = False


class IntVar(Var):
    default = 0


class Widget:
    def __init__(self, *a, **kw):
        self.opts = dict(kw)
        self.store = {}
        self.grid_calls = 0

    def __getattr__(self, name):  # pack/grid/bind/configure/... - заглушки
        if name.startswith("__"):
            raise AttributeError(name)

        def noop(*a, **k):
            if name in ("configure", "config"):
                self.opts.update(k)
            return None

        return noop

    def __setitem__(self, k, v):
        self.store[k] = v

    def __getitem__(self, k):
        return self.store.get(k, 0)


class FakeTk(Widget):
    def __init__(self, *a, **kw):
        super().__init__()
        self.afters = []
        self.destroyed = False

    def after(self, ms, fn=None, *args):
        self.afters.append((ms, fn))
        return len(self.afters)

    def destroy(self):
        self.destroyed = True


class Box:
    """Записує виклики messagebox/filedialog."""

    def __init__(self):
        self.calls = []
        self.answers = {}

    def fn(self, name, default=None):
        def _f(*a, **k):
            self.calls.append((name, a, k))
            return self.answers.get(name, default)

        return _f


@pytest.fixture
def gui(monkeypatch):
    box = Box()
    tk = types.ModuleType("tkinter")
    ttk = types.ModuleType("tkinter.ttk")
    fd = types.ModuleType("tkinter.filedialog")
    mb = types.ModuleType("tkinter.messagebox")
    tk.Tk = FakeTk
    tk.StringVar, tk.BooleanVar, tk.IntVar = StringVar, BooleanVar, IntVar
    tk.BOTH, tk.LEFT = "both", "left"
    for name in ("Frame", "Label", "Entry", "Button", "Checkbutton", "Combobox", "Spinbox", "Progressbar"):
        setattr(ttk, name, Widget)

    class RecordingStyle(Widget):
        maps = {}

        def map(self, style_name, **kw):
            RecordingStyle.maps[style_name] = kw

    RecordingStyle.maps = {}
    ttk.Style = RecordingStyle
    box.style = RecordingStyle
    fd.askdirectory = box.fn("askdirectory", "")
    fd.askopenfilename = box.fn("askopenfilename", "")
    mb.showerror = box.fn("showerror")
    mb.showinfo = box.fn("showinfo")
    mb.askyesno = box.fn("askyesno", True)
    tk.ttk, tk.filedialog, tk.messagebox = ttk, fd, mb
    for name, mod in (("tkinter", tk), ("tkinter.ttk", ttk), ("tkinter.filedialog", fd), ("tkinter.messagebox", mb)):
        monkeypatch.setitem(sys.modules, name, mod)
    sys.modules.pop("DocxImageCompressor", None)
    mod = importlib.import_module("DocxImageCompressor")
    yield mod, box
    sys.modules.pop("DocxImageCompressor", None)


def finish(app):
    """Дочекатись потоку й вичерпати чергу так, як це робить after()."""
    app._thread.join(timeout=30)
    assert not app._thread.is_alive()
    app._poll_queue()


def messages(box, name):
    return [c[1][1] for c in box.calls if c[0] == name]


def test_builds_and_convert_is_default_first_mode(gui):
    mod, box = gui
    app = mod.App()
    assert mod.MODES[0] == mod.MODE_CONVERT
    assert app.mode_var.get() == mod.MODE_CONVERT
    assert app.mode_cb.opts["values"][0] == "Конвертувати .doc → .docx"
    assert list(app.mode_cb.opts["values"]) == list(mod.MODES)


@pytest.mark.parametrize("idx", range(5))
def test_every_mode_can_be_selected(gui, idx):
    mod, box = gui
    app = mod.App()
    app.mode_var.set(mod.MODES[idx])
    app._on_mode_change()  # не має падати


def test_start_requires_input_folder(gui):
    mod, box = gui
    app = mod.App()
    app.start()
    assert messages(box, "showerror") and not app.running


def test_images_run_goes_through_queue_and_shows_result(gui, tmp_path):
    mod, box = gui
    Image.effect_noise((1800, 1200), 60).convert("RGB").save(tmp_path / "a.png")
    app = mod.App()
    app.mode_var.set(mod.MODE_IMAGES)
    app.input_var.set(str(tmp_path))
    app.output_var.set(str(tmp_path / "out"))
    app.start()
    assert app.running and app.cancel_btn.opts["state"] == "normal"
    finish(app)
    assert not app.running and app.cancel_btn.opts["state"] == "disabled"
    result = messages(box, "showinfo")[-1]
    assert "Оброблено: 1/1" in result
    assert (tmp_path / "out" / "a.jpg").exists()


def test_in_place_overwrite_asks_for_confirmation(gui, tmp_path):
    mod, box = gui
    Image.new("RGB", (10, 10)).save(tmp_path / "a.png")
    app = mod.App()
    app.mode_var.set(mod.MODE_IMAGES)
    app.input_var.set(str(tmp_path))
    app.same_folder.set(True)
    box.answers["askyesno"] = False
    app.start()
    assert not app.running and app._thread is None  # користувач відмовився
    box.answers["askyesno"] = True
    app.start()
    finish(app)
    assert any("замінено на місці" in a[1][1] for a in box.calls if a[0] == "askyesno")


def test_exif_without_backup_asks_for_confirmation(gui, tmp_path):
    mod, box = gui
    Image.new("RGB", (10, 10)).save(tmp_path / "a.jpg")
    app = mod.App()
    app.mode_var.set(mod.MODE_EXIF)
    app.input_var.set(str(tmp_path))
    app.backup_var.set(False)
    box.answers["askyesno"] = False
    app.start()
    assert not app.running


def test_rename_invalid_template_blocked_before_start(gui, tmp_path):
    mod, box = gui
    Image.new("RGB", (10, 10)).save(tmp_path / "a.jpg")
    app = mod.App()
    app.mode_var.set(mod.MODE_RENAME)
    app.input_var.set(str(tmp_path))
    app.template_var.set("{bogus}")
    app.start()
    assert messages(box, "showerror") and not app.running


def test_rename_invalid_start_number(gui, tmp_path):
    mod, box = gui
    app = mod.App()
    app.mode_var.set(mod.MODE_RENAME)
    app.input_var.set(str(tmp_path))
    app.start_num_var.set("abc")
    app.start()
    assert messages(box, "showerror") and not app.running


def test_convert_without_pywin32_is_reported(gui, tmp_path, monkeypatch):
    mod, box = gui
    monkeypatch.setattr(mod, "WIN32_OK", False)
    app = mod.App()
    app.input_var.set(str(tmp_path))
    app.output_var.set(str(tmp_path))
    app.start()
    assert any("pywin32" in m for m in messages(box, "showerror"))


def test_cancel_sets_event_and_disables_button(gui, tmp_path):
    mod, box = gui
    app = mod.App()
    app.running = True
    app.cancel()
    assert app._cancel.is_set() and app.cancel_btn.opts["state"] == "disabled"


def test_worker_crash_is_reported_with_log_hint(gui, tmp_path, monkeypatch):
    mod, box = gui
    app = mod.App()

    def boom(*a, **k):
        raise RuntimeError("kaboom")

    monkeypatch.setattr(mod.core, "run_job", boom)
    app.running = True
    app._worker(object())
    app._poll_queue()
    assert "kaboom" in messages(box, "showinfo")[-1]
    assert not app.running


def test_undo_rename_button_flow(gui, tmp_path):
    mod, box = gui
    import techoffice_core as c

    for n in ("a.jpg", "b.jpg"):
        Image.new("RGB", (8, 8)).save(tmp_path / n)
    files = c._find_files(tmp_path, False, c.IMG_EXTS)
    done, _ = c._rename_many(c.plan_renames(files, "N{n}", 1, "За іменем"))
    log = c.write_rename_log(tmp_path, done)
    app = mod.App()
    box.answers["askopenfilename"] = str(log)
    app.undo_rename()
    assert "Повернуто назв: 2" in messages(box, "showinfo")[-1]
    assert sorted(p.name for p in tmp_path.iterdir() if p.suffix == ".jpg") == ["a.jpg", "b.jpg"]


def test_undo_rename_with_bad_log_shows_error(gui, tmp_path):
    mod, box = gui
    bad = tmp_path / "bad.json"
    bad.write_text("not json")
    app = mod.App()
    box.answers["askopenfilename"] = str(bad)
    app.undo_rename()
    assert messages(box, "showerror")


def test_close_when_idle_destroys_window(gui):
    mod, box = gui
    app = mod.App()
    app._on_close()
    assert app.destroyed


def test_close_while_running_asks_and_cancels(gui):
    mod, box = gui
    app = mod.App()
    app.running = True
    box.answers["askyesno"] = False
    app._on_close()
    assert not app._cancel.is_set() and not app.destroyed  # користувач передумав
    box.answers["askyesno"] = True
    app._on_close()
    assert app._cancel.is_set()
    app._wait_and_close()  # потоку немає -> закриваємо
    assert app.destroyed


def test_missing_dependency_shows_message_instead_of_silent_exit(monkeypatch, gui):
    mod, box = gui
    sys.modules.pop("DocxImageCompressor", None)
    monkeypatch.setitem(sys.modules, "techoffice_core", None)  # імітуємо ImportError
    with pytest.raises(SystemExit):
        importlib.import_module("DocxImageCompressor")
    assert any("requirements.txt" in c[1][1] for c in box.calls if c[0] == "showerror")


def test_hover_and_disabled_states_keep_text_readable(gui):
    """Регресія: при наведенні фон ставав світлим, а текст лишався білим."""
    mod, box = gui
    mod.App()
    maps = box.style.maps
    for name in ("TButton", "TCheckbutton", "TEntry", "TCombobox", "TSpinbox"):
        assert name in maps, name
    btn = dict(maps["TButton"])
    states = [st for st, _ in btn["background"]]
    assert "active" in states and "disabled" in states and "pressed" in states
    # світлого "#ececec" з теми clam більше немає ні в одному стані
    for style_name, kw in maps.items():
        for option, pairs in kw.items():
            for state, colour in pairs:
                assert str(colour).lower() not in ("#ececec", "#d9d9d9"), (style_name, option, state)
    # у вимкненому стані текст кнопки не білий (не зливається зі світлим фоном)
    fg = dict((tuple([st]) if isinstance(st, str) else st, c) for st, c in btn["foreground"])
    assert fg[("disabled",)] != "white"


def test_drive_root_with_subfolders_asks_before_starting(gui):
    mod, box = gui
    root = str(Path("/").resolve())  # на Windows - корінь поточного диска
    app = mod.App()
    app.mode_var.set(mod.MODE_RENAME)
    app.input_var.set(root)
    app.recursive_var.set(True)
    box.answers["askyesno"] = False
    app.start()
    assert not app.running and app._thread is None
    assert any("корінь диска" in c[1][1] for c in box.calls if c[0] == "askyesno")


def test_progress_message_during_scan_does_not_break_progressbar(gui):
    mod, box = gui
    app = mod.App()
    app._upd(0, 0, "Пошук файлів… переглянуто 400")
    assert app.status.opts["text"] == "Пошук файлів…"
    assert "400" in app.detail.opts["text"]

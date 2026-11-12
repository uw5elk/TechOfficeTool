"""Логіка конвертації .doc -> .docx перевіряється на підробленому Word
(справжній COM є тільки у Windows)."""
import io
import threading
import zipfile
from pathlib import Path

import pytest

import techoffice_core as c


def minimal_docx_bytes():
    buf = io.BytesIO()
    with zipfile.ZipFile(buf, "w") as z:
        z.writestr("[Content_Types].xml", "<Types/>")
        z.writestr("word/document.xml", "<w:document/>")
    return buf.getvalue()


class FakeDoc:
    def __init__(self, behavior):
        self.behavior = behavior

    def SaveAs(self, FileName, FileFormat):
        assert FileFormat == 16
        if self.behavior == "garbage":
            Path(FileName).write_bytes(b"not a docx")
        elif self.behavior == "empty":
            Path(FileName).write_bytes(b"")
        elif self.behavior == "save_error":
            raise RuntimeError("disk full")
        else:
            Path(FileName).write_bytes(minimal_docx_bytes())

    def Close(self, SaveChanges=False):
        pass


class FakeDocuments:
    def __init__(self, log):
        self.log = log

    def Open(self, **kw):
        name = Path(kw["FileName"]).name
        self.log.append(kw)
        if "locked" in name:
            raise RuntimeError("password protected")
        if "garbage" in name:
            return FakeDoc("garbage")
        if "empty" in name:
            return FakeDoc("empty")
        if "fullsave" in name:
            return FakeDoc("save_error")
        return FakeDoc("ok")


class FakeWord:
    Version = "16.0"

    def __init__(self):
        self.calls = []
        self.Documents = FakeDocuments(self.calls)
        self.quit = False

    def Quit(self):
        self.quit = True


def conv_with(word=None):
    conv = c.WordConverter()
    conv.word = word or FakeWord()
    return conv


class Factory:
    """Замінник WordConverter для run_job."""

    def __init__(self):
        self.conv = conv_with()

    def __call__(self):
        return self

    def __enter__(self):
        return self.conv

    def __exit__(self, *a):
        return False


def make_doc(path, text=b"legacy"):
    path.parent.mkdir(parents=True, exist_ok=True)
    path.write_bytes(text)
    return path


def test_successful_conversion_is_atomic_and_clean(tmp_path):
    src = make_doc(tmp_path / "a.doc")
    r = conv_with().convert(src, src)
    assert r["ok"] and r["out"].name == "a.docx"
    assert c.is_valid_docx_deep(r["out"])
    assert sorted(p.name for p in tmp_path.iterdir()) == ["a.doc", "a.docx"]  # без тимчасових файлів


def test_document_opened_read_only_without_password_prompt_and_macros_disabled(tmp_path):
    w = FakeWord()
    conv = conv_with(w)
    conv.convert(make_doc(tmp_path / "a.doc"), tmp_path / "a.doc")
    kw = w.calls[0]
    assert kw["ReadOnly"] is True and kw["ConfirmConversions"] is False
    assert kw["PasswordDocument"]  # задано, щоб Word не показував діалог пароля


def test_existing_docx_is_not_overwritten(tmp_path):
    src = make_doc(tmp_path / "a.doc")
    existing = tmp_path / "a.docx"
    existing.write_bytes(b"precious")
    r = conv_with().convert(src, src)
    assert existing.read_bytes() == b"precious"
    assert r["out"].name == "a_1.docx"


@pytest.mark.parametrize("name", ["locked.doc", "garbage.doc", "empty.doc", "fullsave.doc"])
def test_failures_leave_no_partial_docx_and_keep_source(tmp_path, name):
    src = make_doc(tmp_path / name)
    r = conv_with().convert(src, src)
    assert r["ok"] is False and r["error"]
    assert src.read_bytes() == b"legacy"
    assert [p.name for p in tmp_path.iterdir()] == [name]


def test_failure_copies_source_into_output_folder(tmp_path):
    src = make_doc(tmp_path / "in" / "locked.doc")
    dst = tmp_path / "out" / "locked.doc"
    r = conv_with().convert(src, dst)
    assert r["ok"] is False
    assert dst.read_bytes() == b"legacy"


def test_dead_word_is_restarted(tmp_path, monkeypatch):
    class Dead:
        @property
        def Version(self):
            raise RuntimeError("RPC server unavailable")

        def Quit(self):
            raise RuntimeError("dead")

    conv = c.WordConverter()
    conv.word = Dead()
    started = []

    def fake_start():
        started.append(1)
        conv.word = FakeWord()

    monkeypatch.setattr(conv, "_start", fake_start)
    r = conv.convert(make_doc(tmp_path / "a.doc"), tmp_path / "a.doc")
    assert started and r["ok"]


# ----------------------------------------------------------------- run_job --
def cfg(tmp, **kw):
    return c.JobConfig(mode=c.MODE_CONVERT, input_dir=tmp, output_dir=kw.pop("out", tmp), **kw)


def test_run_job_convert_reports_failures_with_names(tmp_path):
    make_doc(tmp_path / "ok.doc")
    make_doc(tmp_path / "locked.doc")
    msg = c.run_job(cfg(tmp_path), converter_factory=Factory())
    assert "Конвертовано: 1, Помилок: 1" in msg
    assert "locked.doc" in msg  # раніше цей список завжди був порожнім


def test_run_job_delete_originals_only_after_success(tmp_path, monkeypatch):
    monkeypatch.setattr(c, "_send2trash", None)
    make_doc(tmp_path / "ok.doc")
    make_doc(tmp_path / "locked.doc")
    msg = c.run_job(cfg(tmp_path, delete_doc=True), converter_factory=Factory())
    assert not (tmp_path / "ok.doc").exists() and (tmp_path / "ok.docx").exists()
    assert (tmp_path / "locked.doc").exists()  # невдалий - оригінал на місці
    assert "Видалено оригіналів .doc" in msg


def test_run_job_never_deletes_when_output_is_invalid(tmp_path, monkeypatch):
    monkeypatch.setattr(c, "_send2trash", None)
    make_doc(tmp_path / "garbage.doc")
    c.run_job(cfg(tmp_path, delete_doc=True), converter_factory=Factory())
    assert (tmp_path / "garbage.doc").exists()


def test_run_job_uses_trash_when_available(tmp_path, monkeypatch):
    trashed = []
    monkeypatch.setattr(c, "_send2trash", lambda p: trashed.append(p))
    make_doc(tmp_path / "ok.doc")
    msg = c.run_job(cfg(tmp_path, delete_doc=True), converter_factory=Factory())
    assert trashed and (tmp_path / "ok.doc").exists()  # заглушка нічого не видаляє
    assert "у кошик" in msg


def test_run_job_one_word_instance_for_whole_batch(tmp_path):
    for i in range(5):
        make_doc(tmp_path / f"f{i}.doc")
    f = Factory()
    c.run_job(cfg(tmp_path), converter_factory=f)
    assert len(f.conv.word.calls) == 5


def test_run_job_recursive_flag_and_output_folder(tmp_path):
    make_doc(tmp_path / "top.doc")
    make_doc(tmp_path / "sub" / "deep.doc")
    out = tmp_path / "out"
    c.run_job(cfg(tmp_path, out=out, recursive=True), converter_factory=Factory())
    assert (out / "top.docx").exists() and (out / "sub" / "deep.docx").exists()
    # другий прогін: вже сконвертоване пропускається, дублікатів x_1.docx немає
    msg = c.run_job(cfg(tmp_path, out=out, recursive=True), converter_factory=Factory())
    assert "Конвертовано: 0" in msg and "Пропущено (поруч уже є .docx): 2" in msg
    assert sorted(p.name for p in out.rglob("*.docx")) == ["deep.docx", "top.docx"]


def test_run_job_non_recursive_ignores_subfolders(tmp_path):
    make_doc(tmp_path / "top.doc")
    make_doc(tmp_path / "sub" / "deep.doc")
    msg = c.run_job(cfg(tmp_path), converter_factory=Factory())
    assert "Конвертовано: 1" in msg and not (tmp_path / "sub" / "deep.docx").exists()


def test_run_job_cancel_stops_batch(tmp_path):
    for i in range(4):
        make_doc(tmp_path / f"f{i}.doc")
    ev = threading.Event()

    def prog(i, t, n):
        if i == 2:
            ev.set()

    msg = c.run_job(cfg(tmp_path), progress=prog, cancel=ev, converter_factory=Factory())
    assert "Скасовано користувачем" in msg
    assert len(list(tmp_path.glob("*.docx"))) == 2


def test_run_job_without_pywin32_gives_clear_message(tmp_path, monkeypatch):
    monkeypatch.setattr(c, "WIN32_OK", False)
    make_doc(tmp_path / "a.doc")
    assert "pywin32" in c.run_job(cfg(tmp_path))


def test_no_docs_message(tmp_path):
    assert "не знайдено" in c.run_job(cfg(tmp_path), converter_factory=Factory())


def test_run_job_skips_doc_when_docx_exists_and_never_deletes_it(tmp_path, monkeypatch):
    monkeypatch.setattr(c, "_send2trash", None)
    make_doc(tmp_path / "a.doc")
    (tmp_path / "a.docx").write_bytes(b"user edited version")
    msg = c.run_job(cfg(tmp_path, delete_doc=True), converter_factory=Factory())
    assert "Пропущено (поруч уже є .docx): 1" in msg
    assert (tmp_path / "a.docx").read_bytes() == b"user edited version"
    assert (tmp_path / "a.doc").exists()  # пропущений оригінал не видаляється
    assert not (tmp_path / "a_1.docx").exists()


def test_run_job_word_cannot_start_gives_clear_message(tmp_path):
    make_doc(tmp_path / "a.doc")

    def broken_factory():
        raise RuntimeError("Invalid class string")

    msg = c.run_job(cfg(tmp_path), converter_factory=broken_factory)
    assert "Не вдалося запустити Microsoft Word" in msg and "Invalid class string" in msg
    assert (tmp_path / "a.doc").read_bytes() == b"legacy"


def test_word_converter_cleans_up_com_when_start_fails(monkeypatch):
    conv = c.WordConverter()
    closed = []

    def bad_start():
        conv._com_inited = True
        raise RuntimeError("no Word")

    monkeypatch.setattr(conv, "_start", bad_start)
    monkeypatch.setattr(conv, "close", lambda: closed.append(1))
    with pytest.raises(RuntimeError):
        with conv:
            pass
    assert closed == [1]

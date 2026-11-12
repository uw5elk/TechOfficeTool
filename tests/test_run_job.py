import threading

import pytest
from PIL import Image

import techoffice_core as c

ORIENT = 0x0112


def level(name="Для екрану (≈150 ppi)"):
    return c.LEVELS[name]


def cfg(mode, inp, out=None, **kw):
    return c.JobConfig(mode=mode, input_dir=inp, output_dir=out or inp, **kw)


def test_mode_order_convert_is_first():
    assert c.MODES[0] == c.MODE_CONVERT == "Конвертувати .doc → .docx"
    assert set(c.MODES) == {c.MODE_CONVERT, c.MODE_DOCX, c.MODE_IMAGES, c.MODE_RENAME, c.MODE_EXIF}
    assert len(c.MODES) == 5


def test_images_mode_to_other_folder_leaves_sources(tmp_path, noise):
    src = tmp_path / "in"
    (src / "sub").mkdir(parents=True)
    noise().save(src / "a.png")
    noise().save(src / "sub" / "b.jpg", quality=100)
    before = {p.name: p.read_bytes() for p in src.rglob("*") if p.is_file()}
    out = tmp_path / "out"
    msg = c.run_job(cfg(c.MODE_IMAGES, src, out, level=level(), recursive=True))
    assert "Оброблено: 2/2" in msg and "Зменшення" in msg
    assert {p.name: p.read_bytes() for p in src.rglob("*") if p.is_file()} == before
    assert (out / "a.jpg").exists() and (out / "sub" / "b.jpg").exists()


def test_images_mode_non_recursive_by_default(tmp_path, noise):
    (tmp_path / "sub").mkdir()
    noise().save(tmp_path / "a.jpg", quality=100)
    noise().save(tmp_path / "sub" / "b.jpg", quality=100)
    sub_before = (tmp_path / "sub" / "b.jpg").read_bytes()
    msg = c.run_job(cfg(c.MODE_IMAGES, tmp_path, level=level()))
    assert "Оброблено: 1/1" in msg
    assert (tmp_path / "sub" / "b.jpg").read_bytes() == sub_before


def test_output_folder_inside_input_is_not_reprocessed(tmp_path, noise):
    noise().save(tmp_path / "a.jpg", quality=100)
    out = tmp_path / "compressed"
    c.run_job(cfg(c.MODE_IMAGES, tmp_path, out, level=level(), recursive=True))
    first = (out / "a.jpg").read_bytes()
    msg = c.run_job(cfg(c.MODE_IMAGES, tmp_path, out, level=level(), recursive=True))
    assert "Оброблено: 1/1" in msg  # лише вихідний a.jpg, а не вже стиснений
    # результат детермінований: повторний прогін дає ті самі байти (джерело - оригінал, а не копія)
    assert (out / "a.jpg").read_bytes() == first
    assert sorted(p.name for p in out.iterdir()) == ["a.jpg"]  # і жодних a_1.jpg / вкладених копій


def test_images_mode_reports_corrupt_files_by_name(tmp_path, noise):
    noise().save(tmp_path / "ok.jpg", quality=100)
    (tmp_path / "broken.jpg").write_bytes(b"xx" * 100)
    msg = c.run_job(cfg(c.MODE_IMAGES, tmp_path, level=level()))
    assert "Помилок: 1" in msg and "broken.jpg" in msg
    assert (tmp_path / "broken.jpg").read_bytes() == b"xx" * 100


def test_images_mode_keep_exif_flag(tmp_path, noise):
    ex = Image.Exif()
    ex[0x010F] = "Cam"
    noise().save(tmp_path / "a.jpg", exif=ex.tobytes(), quality=100)
    noise().save(tmp_path / "b.jpg", exif=ex.tobytes(), quality=100)
    c.run_job(cfg(c.MODE_IMAGES, tmp_path, level=level(), keep_exif=True))
    with Image.open(tmp_path / "a.jpg") as im:
        assert im.getexif().get(0x010F) == "Cam"
    c.run_job(cfg(c.MODE_IMAGES, tmp_path, level=level("Сильне стиснення"), keep_exif=False))
    with Image.open(tmp_path / "b.jpg") as im:
        assert dict(im.getexif()) == {}


def test_docx_mode_recursive_and_counts(tmp_path, noise):
    docx = pytest.importorskip("docx")
    from docx.shared import Inches

    (tmp_path / "sub").mkdir()
    img = tmp_path / "p.png"
    noise().save(img)
    for rel in ("one.docx", "sub/two.docx"):
        d = docx.Document()
        d.add_picture(str(img), width=Inches(3))
        d.save(tmp_path / rel)
    (tmp_path / "~$lock.docx").write_bytes(b"PK\x03\x04junk")  # службовий файл Word - ігнорується
    msg = c.run_job(cfg(c.MODE_DOCX, tmp_path, level=level()))
    assert "Файлів: 2, Зображень: 2" in msg and "Зменшення" in msg
    assert (tmp_path / "~$lock.docx").read_bytes() == b"PK\x03\x04junk"


def test_rename_dry_run_changes_nothing_then_real(tmp_path):
    for n in ("b.jpg", "a.jpg"):
        Image.new("RGB", (8, 8)).save(tmp_path / n)
    msg = c.run_job(cfg(c.MODE_RENAME, tmp_path, dry_run=True))
    assert "a.jpg → IMG_0001.jpg" in msg
    assert sorted(p.name for p in tmp_path.iterdir()) == ["a.jpg", "b.jpg"]
    msg = c.run_job(cfg(c.MODE_RENAME, tmp_path, dry_run=False))
    assert "Перейменовано: 2" in msg and "Журнал для відміни" in msg
    logs = list(tmp_path.glob(".rename_log_*.json"))
    assert len(logs) == 1
    restored, skipped, errs = c.undo_rename(logs[0])
    assert restored == 2
    assert sorted(p.name for p in tmp_path.iterdir() if p.suffix == ".jpg") == ["a.jpg", "b.jpg"]


def test_rename_invalid_template_is_reported_not_crashed(tmp_path):
    Image.new("RGB", (8, 8)).save(tmp_path / "a.jpg")
    msg = c.run_job(cfg(c.MODE_RENAME, tmp_path, template="{oops}", dry_run=False))
    assert "шаблон" in msg.lower()
    assert (tmp_path / "a.jpg").exists()


def test_exif_mode_counts_and_backup_warning(tmp_path, noise):
    ex = Image.Exif()
    ex[0x010F] = "Cam"
    noise((300, 200)).save(tmp_path / "dirty.jpg", exif=ex.tobytes())
    noise((300, 200)).save(tmp_path / "clean.jpg")
    msg = c.run_job(cfg(c.MODE_EXIF, tmp_path, backup=True))
    assert "Очищено: 1" in msg and "Без EXIF: 1" in msg and "Помилок: 0" in msg
    assert "*.bak" in msg and "EXIF/GPS" in msg
    assert (tmp_path / "dirty.jpg.bak").exists()


def test_cancel_mid_batch_in_exif_mode(tmp_path, noise):
    ex = Image.Exif()
    ex[0x010F] = "Cam"
    for i in range(4):
        noise((200, 150)).save(tmp_path / f"{i}.jpg", exif=ex.tobytes())
    ev = threading.Event()

    def prog(i, t, n):
        if i == 2:
            ev.set()

    msg = c.run_job(cfg(c.MODE_EXIF, tmp_path, backup=False), progress=prog, cancel=ev)
    assert "Скасовано користувачем" in msg and "Очищено: 2" in msg


def test_progress_callback_receives_totals(tmp_path, noise):
    for i in range(3):
        noise((300, 200)).save(tmp_path / f"{i}.png")
    seen = []
    c.run_job(cfg(c.MODE_IMAGES, tmp_path, level=level()), progress=lambda i, t, n: seen.append((i, t)))
    assert seen == [(1, 3), (2, 3), (3, 3)]


def test_unknown_mode():
    assert "Невідомий режим" in c.run_job(c.JobConfig(mode="x", input_dir=".", output_dir="."))


def test_cancel_during_file_search_stops_scan(tmp_path):
    """Регресія: «Скасувати» не зупиняло пошук файлів (напр. по всьому диску)."""
    for i in range(1000):
        (tmp_path / f"f{i}.txt").write_text("x")
    ev = threading.Event()
    ev.set()  # скасовано ще до початку
    found = c._find_files(tmp_path, True, c.IMG_EXTS, cancel=ev)
    assert found == []  # перервано на першій же перевірці (кожні 200 елементів)
    for mode in (c.MODE_RENAME, c.MODE_EXIF, c.MODE_IMAGES, c.MODE_DOCX, c.MODE_CONVERT):
        msg = c.run_job(
            cfg(mode, tmp_path, level=level(), recursive=True),
            cancel=ev,
            converter_factory=lambda: None,
        )
        assert "під час пошуку" in msg or "pywin32" in msg, (mode, msg)


def test_scan_reports_progress_with_zero_total(tmp_path):
    for i in range(450):
        (tmp_path / f"f{i}.txt").write_text("x")
    seen = []
    c.run_job(cfg(c.MODE_EXIF, tmp_path, recursive=True), progress=lambda i, t, n: seen.append((i, t, n)))
    assert seen and all(i == 0 and t == 0 and "Пошук" in n for i, t, n in seen)

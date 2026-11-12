import json
import threading

import pytest
from PIL import Image

import techoffice_core as c


def mk(folder, *names):
    folder.mkdir(parents=True, exist_ok=True)
    out = []
    for n in names:
        p = folder / n
        Image.new("RGB", (8, 8), "red").save(p)
        out.append(p)
    return out


def run_rename(folder, template="IMG_{n:04d}", start=1, sort="За іменем", recursive=False):
    files = c._find_files(folder, recursive, c.IMG_EXTS)
    planned = c.plan_renames(files, template, start, sort)
    done, errors = c._rename_many(planned)
    return planned, done, errors


def names(folder):
    return sorted(p.name for p in folder.iterdir() if p.is_file() and not p.name.startswith("."))


def test_basic_rename_in_name_order(tmp_path):
    mk(tmp_path, "b.jpg", "a.jpg", "c.png")
    _, done, errors = run_rename(tmp_path)
    assert not errors and len(done) == 3
    assert names(tmp_path) == ["IMG_0001.jpg", "IMG_0002.jpg", "IMG_0003.png"]


def test_rerun_is_idempotent_no_suffixes(tmp_path):
    mk(tmp_path, "a.jpg", "b.jpg", "c.jpg")
    run_rename(tmp_path)
    before = names(tmp_path)
    planned, done, errors = run_rename(tmp_path)
    assert names(tmp_path) == before == ["IMG_0001.jpg", "IMG_0002.jpg", "IMG_0003.jpg"]
    assert done == [] and not errors  # нічого переіменовувати


def test_chain_conflict_does_not_overwrite(tmp_path):
    # IMG_0002 вже існує і сам буде перейменований -> ланцюжок
    mk(tmp_path, "IMG_0002.jpg", "x.jpg", "y.jpg")
    for p, color in zip(sorted(tmp_path.iterdir()), ("red", "green", "blue")):
        Image.new("RGB", (8, 8), color).save(p)
    contents_before = sorted(p.read_bytes() for p in tmp_path.iterdir())
    _, done, errors = run_rename(tmp_path)
    assert not errors
    assert names(tmp_path) == ["IMG_0001.jpg", "IMG_0002.jpg", "IMG_0003.jpg"]
    assert sorted(p.read_bytes() for p in tmp_path.iterdir()) == contents_before  # жоден файл не втрачено


def test_foreign_file_with_target_name_forces_suffix(tmp_path):
    (a,) = mk(tmp_path, "a.jpg")
    foreign = tmp_path / "IMG_0001.jpg"
    foreign.write_bytes(b"keep me")
    # у плані лише a.jpg, тому IMG_0001.jpg - стороннє, зайняте ім'я
    planned = c.plan_renames([a], "IMG_{n:04d}", 1, "За іменем")
    assert planned[0][1].name == "IMG_0001_1.jpg"
    done, errors = c._rename_many(planned)
    assert not errors and foreign.read_bytes() == b"keep me"
    assert (tmp_path / "IMG_0001_1.jpg").exists()


def test_case_only_rename(tmp_path):
    mk(tmp_path, "PHOTO.JPG")
    planned = c.plan_renames([tmp_path / "PHOTO.JPG"], "photo", 1, "Без сортування")
    done, errors = c._rename_many(planned)
    assert not errors
    assert names(tmp_path) == ["photo.JPG"]


def test_template_with_bad_characters_is_sanitized(tmp_path):
    (p,) = mk(tmp_path, "a.jpg")
    planned = c.plan_renames([p], "x/y:{n}", 1, "Без сортування")
    assert planned[0][1].parent == tmp_path  # жодних підпапок
    assert "/" not in planned[0][1].name and ":" not in planned[0][1].name


@pytest.mark.parametrize("tmpl", ["{bogus}", "{n", "", "   ", "{0}"])
def test_invalid_templates_are_rejected(tmpl):
    assert c.validate_template(tmpl)


def test_valid_templates_pass():
    for t in ("IMG_{n:04d}", "{date}_{name}", "{parent}-{n}"):
        assert c.validate_template(t) is None


def test_start_number_and_recursive_numbering(tmp_path):
    mk(tmp_path / "s1", "a.jpg")
    mk(tmp_path / "s2", "b.jpg")
    run_rename(tmp_path, start=10, recursive=True)
    assert names(tmp_path / "s1") == ["IMG_0010.jpg"]
    assert names(tmp_path / "s2") == ["IMG_0011.jpg"]


def test_log_and_undo_roundtrip(tmp_path):
    mk(tmp_path, "b.jpg", "a.jpg")
    original = names(tmp_path)
    planned, done, errors = run_rename(tmp_path)
    log = c.write_rename_log(tmp_path, done)
    assert log and log.name.startswith(".rename_log_") and log.suffix == ".json"
    data = json.loads(log.read_text(encoding="utf-8"))
    assert len(data["renames"]) == 2
    restored, skipped, errs = c.undo_rename(log)
    assert (restored, skipped, errs) == (2, 0, [])
    assert names(tmp_path) == original


def test_undo_of_chain_rename(tmp_path):
    mk(tmp_path, "IMG_0002.jpg", "x.jpg", "y.jpg")
    original = names(tmp_path)
    planned, done, errors = run_rename(tmp_path)
    log = c.write_rename_log(tmp_path, done)
    restored, skipped, errs = c.undo_rename(log)
    assert not errs and restored == len(done)
    assert names(tmp_path) == original


def test_undo_skips_files_that_were_moved_away(tmp_path):
    mk(tmp_path, "a.jpg", "b.jpg")
    planned, done, errors = run_rename(tmp_path)
    log = c.write_rename_log(tmp_path, done)
    (tmp_path / "IMG_0001.jpg").unlink()
    restored, skipped, errs = c.undo_rename(log)
    assert restored == 1 and skipped == 1 and not errs


def test_rename_never_overwrites_existing_target(tmp_path):
    a, b = mk(tmp_path, "a.jpg", "b.jpg")
    b.write_bytes(b"other")
    done, errors = c._rename_many([(a, b)])  # свідомо "погане" завдання
    assert done == [] and errors
    assert b.read_bytes() == b"other" and a.exists()


def test_cancel_in_direct_mode_stops_between_files(tmp_path):
    mk(tmp_path, "a.jpg", "b.jpg", "c.jpg")
    ev = threading.Event()
    files = c._find_files(tmp_path, False, c.IMG_EXTS)
    planned = c.plan_renames(files, "N{n}", 1, "За іменем")

    def prog(i, t, n):
        ev.set()  # скасовуємо одразу після першого файлу

    done, errors = c._rename_many(planned, cancel=ev, progress=prog)
    assert len(done) == 1 and not errors


def test_date_in_template_uses_file_mtime_when_no_exif(tmp_path):
    import os

    (p,) = mk(tmp_path, "a.jpg")
    os.utime(p, (1_600_000_000, 1_600_000_000))  # 2020-09-13
    assert c.format_name("{date}_{n}", 7, p) == "2020-09-13_7"

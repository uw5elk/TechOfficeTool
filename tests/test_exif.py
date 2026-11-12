import io

from PIL import Image, ImageCms
from PIL.PngImagePlugin import PngInfo

import techoffice_core as c

ORIENT = 0x0112


def _exif(orientation=None):
    ex = Image.Exif()
    if orientation:
        ex[ORIENT] = orientation
    ex[0x010F] = "SecretCam"
    ex[0x8825] = {1: "N", 2: (50.0, 27.0, 0.0)}
    return ex.tobytes()


def _pixels(path):
    with Image.open(path) as im:
        im.load()
        return im.mode, im.size, im.tobytes()


def test_lossless_strip_keeps_pixels_and_removes_gps_and_comment(tmp_path, noise):
    p = tmp_path / "a.jpg"
    noise((800, 600)).save(p, exif=_exif(), comment=b"top-secret-comment", quality=92)
    before = _pixels(p)
    assert b"SecretCam" in p.read_bytes()
    assert c.strip_exif_one(p, True, False, False) == "cleaned"
    raw = p.read_bytes()
    assert b"SecretCam" not in raw and b"top-secret-comment" not in raw
    assert _pixels(p) == before  # жодного перекодування пікселів
    with Image.open(p) as im:
        assert dict(im.getexif()) == {}


def test_progressive_jpeg_lossless(tmp_path, noise):
    p = tmp_path / "prog.jpg"
    noise((700, 500)).save(p, exif=_exif(), progressive=True, quality=90)
    before = _pixels(p)
    assert c.strip_exif_one(p, True, False, False) == "cleaned"
    assert _pixels(p) == before


def test_cmyk_jpeg_keeps_adobe_marker_and_pixels(tmp_path):
    p = tmp_path / "cmyk.jpg"
    Image.effect_noise((300, 200), 50).convert("CMYK").save(p, exif=_exif(), quality=90)
    before = _pixels(p)
    assert c.strip_exif_one(p, True, False, False) == "cleaned"
    assert _pixels(p) == before
    assert b"Adobe" in p.read_bytes()


def test_orientation_survives_lossless_strip(tmp_path, noise):
    p = tmp_path / "o.jpg"
    noise((800, 600)).save(p, exif=_exif(6), quality=92)
    assert c.strip_exif_one(p, True, False, False) == "cleaned"
    with Image.open(p) as im:
        assert dict(im.getexif()) == {ORIENT: 6}
        assert im.size == (800, 600)
    # і повторний запуск нічого не змінює
    assert c.strip_exif_one(p, True, False, False) == "already_clean"


def test_paranoid_rotates_pixels_and_removes_everything(tmp_path, noise):
    p = tmp_path / "o.jpg"
    noise((800, 600)).save(p, exif=_exif(6), quality=92)
    assert c.strip_exif_one(p, True, True, False) == "cleaned"
    with Image.open(p) as im:
        assert dict(im.getexif()) == {}
        assert im.size == (600, 800)  # орієнтація "вшита" в пікселі
    assert c.strip_exif_one(p, True, True, False) == "already_clean"


def test_clean_file_is_not_modified(tmp_path, noise):
    p = tmp_path / "clean.jpg"
    noise((200, 100)).save(p)
    before = p.read_bytes()
    assert c.strip_exif_one(p, True, False, True) == "already_clean"
    assert p.read_bytes() == before
    assert not (tmp_path / "clean.jpg.bak").exists()


def test_backup_created_only_when_requested(tmp_path, noise):
    a, b = tmp_path / "a.jpg", tmp_path / "b.jpg"
    for f in (a, b):
        noise((200, 100)).save(f, exif=_exif())
    original_a = a.read_bytes()
    assert c.strip_exif_one(a, True, False, True) == "cleaned"
    assert c.strip_exif_one(b, True, False, False) == "cleaned"
    assert (tmp_path / "a.jpg.bak").read_bytes() == original_a
    assert not (tmp_path / "b.jpg.bak").exists()


def test_png_text_chunks_and_exif_removed(tmp_path, noise):
    p = tmp_path / "t.png"
    meta = PngInfo()
    meta.add_text("Author", "Богдан-secret")
    noise((300, 200)).save(p, pnginfo=meta, exif=_exif())
    assert c.strip_exif_one(p, True, False, False) == "cleaned"
    raw = p.read_bytes()
    assert b"secret" not in raw and b"SecretCam" not in raw
    with Image.open(p) as im:
        assert not getattr(im, "text", {}) and dict(im.getexif()) == {}
    assert c.strip_exif_one(p, True, False, False) == "already_clean"


def test_jpeg_with_trailing_data_falls_back_to_repack(tmp_path, noise):
    p = tmp_path / "mpo_like.jpg"
    noise((300, 200)).save(p, exif=_exif(), quality=90)
    p.write_bytes(p.read_bytes() + b"\xff\xd8\xff\xe1" + b"SecretCam" * 20 + b"\xff\xd9")
    assert c.strip_exif_one(p, True, False, False) == "cleaned"
    assert b"SecretCam" not in p.read_bytes()
    with Image.open(p) as im:
        im.load()
        assert im.size == (300, 200)


def test_icc_kept_or_dropped_on_request(tmp_path, noise):
    icc = ImageCms.ImageCmsProfile(ImageCms.createProfile("sRGB")).tobytes()
    keep, drop = tmp_path / "keep.jpg", tmp_path / "drop.jpg"
    for f in (keep, drop):
        noise((300, 200)).save(f, exif=_exif(), icc_profile=icc)
    assert c.strip_exif_one(keep, True, False, False) == "cleaned"
    assert c.strip_exif_one(drop, False, False, False) == "cleaned"
    with Image.open(keep) as im:
        assert im.info.get("icc_profile")
    with Image.open(drop) as im:
        assert not im.info.get("icc_profile")


def test_error_leaves_file_and_dir_untouched(tmp_path):
    p = tmp_path / "bad.jpg"
    p.write_bytes(b"garbage" * 100)
    before = p.read_bytes()
    assert c.strip_exif_one(p, True, False, True) == "error"
    assert p.read_bytes() == before
    assert [x.name for x in tmp_path.iterdir()] == ["bad.jpg"]


def test_mtime_preserved(tmp_path, noise):
    import os

    p = tmp_path / "a.jpg"
    noise((200, 100)).save(p, exif=_exif())
    os.utime(p, (1_400_000_000, 1_400_000_000))
    c.strip_exif_one(p, True, False, False)
    assert int(p.stat().st_mtime) == 1_400_000_000


def test_strip_jpeg_lossless_rejects_non_jpeg():
    assert c.strip_jpeg_lossless(b"not a jpeg", True) is None


def test_webp_exif_is_removed(tmp_path, noise):
    p = tmp_path / "w.webp"
    noise((300, 200)).save(p, exif=_exif(), quality=90)
    with Image.open(p) as im:
        assert im.getexif()  # передумова: EXIF справді є
    assert c.strip_exif_one(p, True, False, False) == "cleaned"
    with Image.open(p) as im:
        assert dict(im.getexif()) == {}
    assert c.strip_exif_one(p, True, False, False) == "already_clean"


def test_animated_gif_is_skipped_and_untouched(tmp_path):
    frames = [Image.new("RGB", (32, 32), col).convert("P") for col in ("red", "green", "blue")]
    p = tmp_path / "a.gif"
    frames[0].save(p, save_all=True, append_images=frames[1:], duration=40, loop=0, comment=b"secret")
    before = p.read_bytes()
    assert c.strip_exif_one(p, True, False, True) == "skipped"
    assert p.read_bytes() == before and not (tmp_path / "a.gif.bak").exists()


def test_tiff_is_skipped_instead_of_failing(tmp_path, noise):
    p = tmp_path / "s.tif"
    noise((100, 80)).save(p)
    before = p.read_bytes()
    assert c.strip_exif_one(p, True, False, False) == "skipped"
    assert p.read_bytes() == before


def test_animated_png_is_skipped(tmp_path):
    frames = [Image.new("RGB", (32, 32), col) for col in ("red", "green", "blue")]
    p = tmp_path / "a.png"
    frames[0].save(p, save_all=True, append_images=frames[1:], duration=40, loop=0)
    with Image.open(p) as chk:
        assert chk.n_frames == 3
    before = p.read_bytes()
    assert c.strip_exif_one(p, True, True, False) == "skipped"
    assert p.read_bytes() == before

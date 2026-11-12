import os

from PIL import Image, ImageCms

import techoffice_core as c

ORIENT = 0x0112


def _exif(orientation=None, gps=False):
    ex = Image.Exif()
    if orientation:
        ex[ORIENT] = orientation
    ex[0x010F] = "TestCam"
    if gps:
        ex[0x8825] = {1: "N"}
    return ex.tobytes()


def test_png_opaque_becomes_real_jpeg_with_jpg_extension(tmp_path, noise):
    p = tmp_path / "a.png"
    noise().save(p)
    r = c.compress_file(p, p, 1600, 82, True)
    assert r["ok"] and r["status"] == "compressed"
    assert r["path"].name == "a.jpg"
    assert sorted(x.name for x in tmp_path.iterdir()) == ["a.jpg"]  # старий .png замінено
    with Image.open(tmp_path / "a.jpg") as im:
        assert im.format == "JPEG"
        assert max(im.size) == 1600
    assert r["new"] < r["old"]


def test_transparent_png_stays_png(tmp_path, noise):
    img = noise(mode="RGBA")
    img.putalpha(Image.linear_gradient("L").resize(img.size))
    p = tmp_path / "t.png"
    img.save(p)
    r = c.compress_file(p, p, 1000, 82, True)
    assert r["ok"] and r["path"].name == "t.png"
    with Image.open(p) as im:
        assert im.format == "PNG" and im.mode == "RGBA"
        assert max(im.size) == 1000


def test_animated_gif_is_left_untouched(tmp_path):
    frames = [Image.new("RGB", (64, 64), col).convert("P") for col in ("red", "green", "blue")]
    p = tmp_path / "anim.gif"
    frames[0].save(p, save_all=True, append_images=frames[1:], duration=50, loop=0)
    with Image.open(p) as chk:
        assert chk.n_frames == 3  # тест має справді перевіряти анімацію
    before = p.read_bytes()
    r = c.compress_file(p, p, 1600, 82, True)
    assert r["ok"] and r["status"] == "skipped"
    assert p.read_bytes() == before


def test_static_gif_and_bmp_get_correct_extension(tmp_path, noise):
    g = tmp_path / "x.bmp"
    noise((900, 700)).save(g)
    r = c.compress_file(g, g, 1600, 82, True)
    assert r["path"].name == "x.jpg"
    with Image.open(r["path"]) as im:
        assert im.format == "JPEG"


def test_original_kept_when_result_not_smaller(tmp_path, noise):
    p = tmp_path / "small.jpg"
    noise((300, 200)).save(p, quality=30)
    before = p.read_bytes()
    r = c.compress_file(p, p, 1600, 95, True)
    assert r["ok"] and r["status"] == "kept"
    assert p.read_bytes() == before


def test_corrupt_file_is_untouched_and_reported(tmp_path):
    p = tmp_path / "bad.jpg"
    p.write_bytes(b"this is not an image" * 50)
    before = p.read_bytes()
    r = c.compress_file(p, p, 1600, 82, True)
    assert r["ok"] is False
    assert p.read_bytes() == before
    assert [x.name for x in tmp_path.iterdir()] == ["bad.jpg"]  # без тимчасового сміття


def test_ext_change_never_overwrites_existing_file(tmp_path, noise):
    p = tmp_path / "a.png"
    noise().save(p)
    other = tmp_path / "a.jpg"
    other.write_bytes(b"precious")
    r = c.compress_file(p, p, 1600, 82, True)
    assert other.read_bytes() == b"precious"
    assert r["path"].name == "a_1.jpg"


def test_output_to_other_folder_keeps_source(tmp_path, noise):
    src = tmp_path / "in" / "a.png"
    src.parent.mkdir()
    noise().save(src)
    before = src.read_bytes()
    dst = tmp_path / "out" / "a.png"
    r = c.compress_file(src, dst, 1600, 82, True)
    assert src.read_bytes() == before
    assert r["path"].name == "a.jpg" and r["path"].parent.name == "out"


def test_output_copy_made_even_when_original_is_kept(tmp_path, noise):
    src = tmp_path / "in" / "s.jpg"
    src.parent.mkdir()
    noise((300, 200)).save(src, quality=30)
    dst = tmp_path / "out" / "s.jpg"
    r = c.compress_file(src, dst, 1600, 95, True)
    assert r["status"] == "kept"
    assert dst.read_bytes() == src.read_bytes()


def test_mtime_is_preserved(tmp_path, noise):
    p = tmp_path / "a.jpg"
    noise().save(p, quality=100)
    os.utime(p, (1_500_000_000, 1_500_000_000))
    c.compress_file(p, p, 1000, 70, True)
    assert int(p.stat().st_mtime) == 1_500_000_000


def test_exif_keep_mode_preserves_orientation_tag(tmp_path, noise):
    p = tmp_path / "o.jpg"
    noise().save(p, exif=_exif(6, gps=True), quality=95)
    r = c.compress_file(p, p, 1000, 80, True, "keep")
    assert r["status"] == "compressed"
    with Image.open(p) as im:
        assert im.getexif().get(ORIENT) == 6
        assert im.size[0] > im.size[1]  # пікселі НЕ повертались - лише тег


def test_orientation_mode_keeps_only_orientation(tmp_path, noise):
    p = tmp_path / "o.jpg"
    noise().save(p, exif=_exif(6, gps=True), quality=95)
    c.compress_file(p, p, 1000, 80, True, "orientation")
    with Image.open(p) as im:
        assert dict(im.getexif()) == {ORIENT: 6}


def test_strip_mode_rotates_pixels_and_drops_exif(tmp_path, noise):
    p = tmp_path / "o.jpg"
    noise().save(p, exif=_exif(6, gps=True), quality=95)
    c.compress_file(p, p, 1600, 80, True, "strip")
    with Image.open(p) as im:
        assert dict(im.getexif()) == {}
        assert im.size[1] > im.size[0]  # було 1800x1200 + orientation 6 -> стало портретним


def test_icc_profile_is_preserved(tmp_path, noise):
    icc = ImageCms.ImageCmsProfile(ImageCms.createProfile("sRGB")).tobytes()
    p = tmp_path / "c.jpg"
    noise().save(p, icc_profile=icc, quality=98)
    c.compress_file(p, p, 1000, 80, True, "strip")
    with Image.open(p) as im:
        assert im.info.get("icc_profile")


def test_grayscale_stays_grayscale(tmp_path, noise):
    p = tmp_path / "g.png"
    Image.effect_noise((1500, 1000), 30).save(p)  # режим L
    r = c.compress_file(p, p, 1000, 80, False)
    with Image.open(r["path"]) as im:
        assert im.mode == "L"


def test_sixteen_bit_png_is_not_touched(tmp_path):
    p = tmp_path / "deep.png"
    Image.new("I;16", (50, 50), 40000).save(p)
    before = p.read_bytes()
    r = c.compress_file(p, p, 1600, 82, True)
    assert r["ok"] and r["status"] == "skipped"
    assert p.read_bytes() == before


def test_process_one_image_wrapper(tmp_path, noise):
    p = tmp_path / "w.png"
    noise().save(p)
    r = c.process_one_image(p, p, 1600, 82, True, True)
    assert r["ok"] and r["new"] < r["old"]


def _run_images(src, out, recursive=False):
    lv = c.LEVELS["Для екрану (≈150 ppi)"]
    return c.run_job(c.JobConfig(mode=c.MODE_IMAGES, input_dir=src, output_dir=out, level=lv, recursive=recursive))


def test_same_stem_different_formats_do_not_overwrite_each_other_in_output(tmp_path, noise):
    """Регресія: a.bmp -> out/a.jpg, потім a.jpg -> out/a.jpg затирав результат."""
    src, out = tmp_path / "in", tmp_path / "out"
    src.mkdir()
    noise((900, 600)).save(src / "a.bmp")
    noise((900, 600)).save(src / "a.jpg", quality=100)
    noise((900, 600)).save(src / "a.gif")
    msg = _run_images(src, out)
    assert "Оброблено: 3/3" in msg
    produced = sorted(p.name for p in out.iterdir())
    assert len(produced) == 3, produced
    assert "a.jpg" in produced  # "рідне" ім'я лишилось за справжнім a.jpg
    sizes = {p.stat().st_size for p in out.iterdir()}
    assert len(sizes) == 3  # три різні файли, а не копії одного


def test_rerun_to_same_output_folder_does_not_create_duplicates(tmp_path, noise):
    src, out = tmp_path / "in", tmp_path / "out"
    src.mkdir()
    noise().save(src / "a.png")
    noise().save(src / "b.bmp")
    _run_images(src, out)
    first = sorted(p.name for p in out.iterdir())
    _run_images(src, out)
    assert sorted(p.name for p in out.iterdir()) == first == ["a.jpg", "b.jpg"]


def test_in_place_conversion_still_never_overwrites_real_files(tmp_path):
    Image.new("RGB", (900, 600), (220, 20, 20)).save(tmp_path / "a.bmp")  # червоний
    Image.new("RGB", (900, 600), (20, 20, 220)).save(tmp_path / "a.jpg", quality=100)  # синій
    _run_images(tmp_path, tmp_path)
    assert sorted(p.name for p in tmp_path.iterdir()) == ["a.jpg", "a_1.jpg"]

    def colour(name):
        with Image.open(tmp_path / name) as im:
            return im.convert("RGB").getpixel((450, 300))

    assert colour("a.jpg")[2] > 150  # справжній a.jpg (синій) лишився під своїм ім'ям
    assert colour("a_1.jpg")[0] > 150  # колишній a.bmp (червоний) отримав нове ім'я


def test_truncated_jpeg_is_reported_and_left_untouched(tmp_path, noise):
    p = tmp_path / "cut.jpg"
    noise((1800, 1200)).save(p, quality=95)
    p.write_bytes(p.read_bytes()[: p.stat().st_size // 2])  # обрізаний файл
    before = p.read_bytes()
    r = c.compress_file(p, p, 1000, 80, True)
    assert r["ok"] is False
    assert p.read_bytes() == before

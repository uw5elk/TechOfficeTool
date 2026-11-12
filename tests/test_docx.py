import io
import zipfile

import pytest
from PIL import Image

import techoffice_core as c

docx = pytest.importorskip("docx")
from docx.shared import Inches  # noqa: E402


def _png(path, noise, size=(1800, 1200), mode="RGB"):
    noise(size, mode).save(path)
    return path


def _jpg(path, noise, size=(1800, 1200)):
    noise(size).save(path, quality=100)
    return path


def make_docx(path, items, header_image=None):
    d = docx.Document()
    d.add_paragraph("Привіт, світ. Текст документа має лишитись незмінним.")
    for img in items:
        d.add_picture(str(img), width=Inches(4))
    if header_image is not None:
        d.sections[0].header.paragraphs[0].add_run().add_picture(str(header_image), width=Inches(1))
    d.save(path)


def _blobs(path):
    d = docx.Document(str(path))
    out = []
    for rel in d.part.rels.values():
        if "image" in rel.reltype:
            out.append(rel.target_part.blob)
    return d, out


def test_shrinks_and_document_still_opens(tmp_path, noise):
    big_png = _png(tmp_path / "p.png", noise)
    big_jpg = _jpg(tmp_path / "j.jpg", noise)
    f = tmp_path / "doc.docx"
    make_docx(f, [big_png, big_jpg])
    old = f.stat().st_size
    r = c.process_one_docx(f, f, 1600, 82, True)
    assert r["ok"] and r["images"] == 2
    assert r["new"] < old and f.stat().st_size == r["new"]

    d, blobs = _blobs(f)
    assert d.paragraphs[0].text.startswith("Привіт, світ")
    assert len(d.inline_shapes) == 2
    assert len(blobs) == 2
    for b in blobs:
        with Image.open(io.BytesIO(b)) as im:
            im.load()
            assert max(im.size) <= 1600
    # ширина зображення в документі не змінилась (змінились лише пікселі)
    assert all(abs(s.width.inches - 4) < 0.01 for s in d.inline_shapes)


def test_png_converted_to_jpeg_is_renamed_everywhere(tmp_path, noise):
    f = tmp_path / "doc.docx"
    make_docx(f, [_png(tmp_path / "p.png", noise)])
    c.process_one_docx(f, f, 1600, 82, True)
    with zipfile.ZipFile(f) as z:
        names = z.namelist()
        media = [n for n in names if n.startswith("word/media/")]
        assert media and all(n.endswith(".jpeg") or n.endswith(".jpg") for n in media), media
        rels = z.read("word/_rels/document.xml.rels").decode()
        for m in media:
            assert m.split("/")[-1] in rels
        assert "image1.png" not in rels
        ct = z.read("[Content_Types].xml").decode().lower()
        assert 'extension="jpeg"' in ct or 'extension="jpg"' in ct
        assert z.testzip() is None
        assert names[0] == "[Content_Types].xml"  # порядок OPC збережено


def test_header_image_relationships_are_rewritten(tmp_path, noise):
    f = tmp_path / "doc.docx"
    body = _png(tmp_path / "b.png", noise)
    head = tmp_path / "h.png"
    noise((1800, 1200)).save(head)
    make_docx(f, [body], header_image=head)
    c.process_one_docx(f, f, 1600, 82, True)
    d = docx.Document(str(f))
    hdr_blobs = [
        rel.target_part.blob
        for rel in d.sections[0].header.part.rels.values()
        if "image" in rel.reltype
    ]
    assert hdr_blobs
    with Image.open(io.BytesIO(hdr_blobs[0])) as im:
        im.load()


def test_nothing_to_shrink_leaves_file_byte_identical(tmp_path, noise):
    small = tmp_path / "s.jpg"
    noise((200, 150)).save(small, quality=40)
    f = tmp_path / "doc.docx"
    make_docx(f, [small])
    before = f.read_bytes()
    r = c.process_one_docx(f, f, 1600, 95, True)
    assert r["ok"] and r["images"] == 0
    assert f.read_bytes() == before


def test_no_temp_files_left_and_output_folder_variant(tmp_path, noise):
    src = tmp_path / "in" / "doc.docx"
    src.parent.mkdir()
    make_docx(src, [_png(tmp_path / "p.png", noise)])
    before = src.read_bytes()
    dst = tmp_path / "out" / "doc.docx"
    r = c.process_one_docx(src, dst, 1600, 82, True)
    assert r["ok"] and src.read_bytes() == before and dst.exists()
    assert sorted(x.name for x in dst.parent.iterdir()) == ["doc.docx"]
    assert sorted(x.name for x in src.parent.iterdir()) == ["doc.docx"]


def test_corrupt_docx_is_preserved_and_reported(tmp_path):
    f = tmp_path / "bad.docx"
    f.write_bytes(b"PK\x03\x04" + b"garbage" * 100)
    before = f.read_bytes()
    r = c.process_one_docx(f, f, 1600, 82, True)
    assert r["ok"] is False
    assert f.read_bytes() == before
    assert [x.name for x in tmp_path.iterdir()] == ["bad.docx"]


def test_corrupt_docx_copied_to_output_folder(tmp_path):
    f = tmp_path / "in" / "bad.docx"
    f.parent.mkdir()
    f.write_bytes(b"PK\x03\x04" + b"garbage" * 100)
    dst = tmp_path / "out" / "bad.docx"
    r = c.process_one_docx(f, dst, 1600, 82, True)
    assert r["ok"] is False and dst.read_bytes() == f.read_bytes()


def test_unreadable_image_inside_docx_is_skipped_not_fatal(tmp_path, noise):
    f = tmp_path / "doc.docx"
    make_docx(f, [_png(tmp_path / "p.png", noise)])
    # додаємо в медіа "биту" картинку
    with zipfile.ZipFile(f, "a") as z:
        z.writestr("word/media/broken.png", b"not a png")
    r = c.process_one_docx(f, f, 1600, 82, True)
    assert r["ok"] and r["images"] == 1
    with zipfile.ZipFile(f) as z:
        assert z.read("word/media/broken.png") == b"not a png"


def test_media_name_collision_is_resolved(tmp_path, noise):
    f = tmp_path / "doc.docx"
    make_docx(f, [_png(tmp_path / "p.png", noise)])
    with zipfile.ZipFile(f, "a") as z:  # вже існує файл з майбутнім ім'ям
        z.writestr("word/media/image1.jpg", b"unrelated")
    r = c.process_one_docx(f, f, 1600, 82, True)
    assert r["ok"]
    with zipfile.ZipFile(f) as z:
        assert z.read("word/media/image1.jpg") == b"unrelated"
        assert z.testzip() is None
    docx.Document(str(f))  # відкривається


def test_transparent_png_keeps_png_format_inside_docx(tmp_path, noise):
    img = noise((1800, 1200), "RGBA")
    img.putalpha(Image.linear_gradient("L").resize(img.size))
    p = tmp_path / "t.png"
    img.save(p)
    f = tmp_path / "doc.docx"
    make_docx(f, [p])
    r = c.process_one_docx(f, f, 1000, 82, True)
    assert r["ok"] and r["images"] == 1
    with zipfile.ZipFile(f) as z:
        media = [n for n in z.namelist() if n.startswith("word/media/")]
        assert media == ["word/media/image1.png"]
        with Image.open(io.BytesIO(z.read(media[0]))) as im:
            assert im.format == "PNG" and im.mode == "RGBA"


# ------------------------------------------------- розмір за ppi на сторінці --
def _media_sizes(path):
    out = {}
    with zipfile.ZipFile(path) as z:
        for n in z.namelist():
            if n.startswith("word/media/"):
                with Image.open(io.BytesIO(z.read(n))) as im:
                    out[n] = im.size
    return out


def _add_vml_picture(d, image_path, style):
    from docx.oxml import parse_xml

    rid, _ = d.part.get_or_add_image(str(image_path))
    xml = (
        '<w:pict xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main" '
        'xmlns:v="urn:schemas-microsoft-com:vml" '
        'xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">'
        '<v:shape %s><v:imagedata r:id="%s"/></v:shape></w:pict>' % (style, rid)
    )
    d.add_paragraph().add_run()._r.append(parse_xml(xml))


def test_small_picture_on_page_is_downscaled_to_ppi(tmp_path, noise):
    img = _png(tmp_path / "p.png", noise)  # 1800x1200
    f = tmp_path / "doc.docx"
    d = docx.Document()
    d.add_picture(str(img), width=Inches(2))
    d.save(f)
    r = c.process_one_docx(f, f, 1600, 82, True, ppi=150)
    assert r["ok"] and r["images"] == 1
    (size,) = _media_sizes(f).values()
    assert 298 <= max(size) <= 304, size  # 2 дюйми * 150 ppi = 300 px
    shapes = docx.Document(str(f)).inline_shapes
    assert abs(shapes[0].width.inches - 2) < 0.01  # на сторінці розмір той самий


def test_without_ppi_only_max_side_applies(tmp_path, noise):
    f = tmp_path / "doc.docx"
    d = docx.Document()
    d.add_picture(str(_png(tmp_path / "p.png", noise)), width=Inches(2))
    d.save(f)
    c.process_one_docx(f, f, 1600, 82, True)
    (size,) = _media_sizes(f).values()
    assert max(size) == 1600


def test_picture_used_twice_keeps_pixels_for_the_larger_use(tmp_path, noise):
    img = _png(tmp_path / "p.png", noise)
    f = tmp_path / "doc.docx"
    d = docx.Document()
    d.add_picture(str(img), width=Inches(2))
    d.add_picture(str(img), width=Inches(5))
    d.save(f)
    c.process_one_docx(f, f, 1600, 82, True, ppi=150)
    sizes = _media_sizes(f)
    assert len(sizes) == 1
    assert 748 <= max(next(iter(sizes.values()))) <= 754  # 5 дюймів * 150 ppi


def test_cropped_picture_keeps_enough_pixels(tmp_path, noise):
    from docx.oxml import OxmlElement
    from docx.oxml.ns import qn

    f = tmp_path / "doc.docx"
    d = docx.Document()
    d.add_picture(str(_png(tmp_path / "p.png", noise)), width=Inches(2))
    blip = d.element.body.xpath(".//a:blip")[0]
    crop = OxmlElement("a:srcRect")
    crop.set("l", "50000")  # видно лише праву половину зображення
    blip.addnext(crop)
    assert blip.getparent().tag == qn("pic:blipFill")
    d.save(f)
    c.process_one_docx(f, f, 1600, 82, True, ppi=150)
    (size,) = _media_sizes(f).values()
    assert 598 <= max(size) <= 604  # видима половина = 2 дюйми -> усе зображення 600 px


def test_reference_with_unknown_size_falls_back_to_max_side(tmp_path, noise):
    img = _png(tmp_path / "p.png", noise)
    f = tmp_path / "doc.docx"
    d = docx.Document()
    d.add_picture(str(img), width=Inches(2))
    _add_vml_picture(d, img, "")  # те саме зображення, але розмір не вказано
    d.save(f)
    c.process_one_docx(f, f, 1600, 82, True, ppi=150)
    (size,) = _media_sizes(f).values()
    assert max(size) == 1600  # не ризикуємо: лише загальне обмеження


def test_vml_picture_size_is_respected(tmp_path, noise):
    img = _png(tmp_path / "p.png", noise)
    f = tmp_path / "doc.docx"
    d = docx.Document()
    _add_vml_picture(d, img, 'style="position:absolute;width:288pt;height:192pt"')  # 4 x 2.67 дюйма
    d.save(f)
    r = c.process_one_docx(f, f, 1600, 82, True, ppi=150)
    assert r["ok"]
    (size,) = _media_sizes(f).values()
    assert 598 <= max(size) <= 604  # 4 дюйми * 150 ppi


def test_ppi_never_upscales_and_respects_max_side(tmp_path, noise):
    f = tmp_path / "doc.docx"
    d = docx.Document()
    d.add_picture(str(_jpg(tmp_path / "j.jpg", noise, (3000, 2000))), width=Inches(40))  # величезний на сторінці
    d.save(f)
    c.process_one_docx(f, f, 1600, 82, True, ppi=220)
    (size,) = _media_sizes(f).values()
    assert max(size) == 1600


def test_run_job_uses_level_ppi_for_docx(tmp_path, noise):
    d = docx.Document()
    d.add_picture(str(_png(tmp_path / "p.png", noise)), width=Inches(2))
    (tmp_path / "docs").mkdir()
    d.save(tmp_path / "docs" / "a.docx")
    lv = c.LEVELS["Для екрану (≈150 ppi)"]
    c.run_job(c.JobConfig(mode=c.MODE_DOCX, input_dir=tmp_path / "docs", output_dir=tmp_path / "docs", level=lv))
    (size,) = _media_sizes(tmp_path / "docs" / "a.docx").values()
    assert 298 <= max(size) <= 304
    assert c.LEVELS["Максимальна якість (архів)"]["ppi"] is None

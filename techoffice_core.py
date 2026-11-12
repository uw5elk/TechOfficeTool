"""TechOfficeTool - ядро: уся логіка без GUI (tkinter сюди не імпортується).

Принципи, на яких побудовано модуль:
  * оригінали ніколи не перезаписуються "напряму": результат спочатку пишеться
    у тимчасовий файл, перевіряється і лише потім атомарно підміняє оригінал;
  * якщо щось пішло не так - оригінал лишається недоторканим, а причина
    потрапляє в журнал (див. setup_logging);
  * якщо "стиснений" файл не менший за оригінал - лишається оригінал.
"""
from __future__ import annotations

import io
import logging
import logging.handlers
import math
import os
import posixpath
import re
import shutil
import struct
import tempfile
import time
import uuid
import xml.etree.ElementTree as ET
import zipfile
from dataclasses import dataclass
from datetime import datetime
from pathlib import Path, PurePosixPath
from typing import Callable, Optional

from PIL import Image, ImageOps

# LOAD_TRUNCATED_IMAGES свідомо НЕ вмикаємо: інакше пошкоджене (обрізане) фото
# "домальовувалось" би сірим і цей результат назавжди замінив би оригінал.

try:  # Word-конвертація доступна лише у Windows з встановленим pywin32
    import pythoncom
    import win32com.client

    WIN32_OK = True
except ImportError:  # pragma: no cover - залежить від платформи
    pythoncom = None
    win32com = None
    WIN32_OK = False

try:  # необов'язково: видалення оригіналів .doc у кошик
    from send2trash import send2trash as _send2trash
except ImportError:
    _send2trash = None

# --------------------------------------------------------------------------- #
# Налаштування
# --------------------------------------------------------------------------- #

LEVELS = {
    "Максимальна якість (архів)": {"max_side": 4000, "quality": 95, "png2jpg": False, "ppi": None},
    "Для друку (≈220 ppi)": {"max_side": 2400, "quality": 88, "png2jpg": True, "ppi": 220},
    "Для екрану (≈150 ppi)": {"max_side": 1600, "quality": 82, "png2jpg": True, "ppi": 150},
    "Сильне стиснення": {"max_side": 1200, "quality": 75, "png2jpg": True, "ppi": 96},
}
# ppi діє лише в .docx: там відомо, якого розміру картинка на сторінці.

MODE_CONVERT = "Конвертувати .doc → .docx"
MODE_DOCX = "Стиснути зображення в .docx"
MODE_IMAGES = "Стиснути звичайні зображення"
MODE_RENAME = "Перейменувати фото"
MODE_EXIF = "Видалити EXIF"
# Порядок = порядок у випадаючому списку. Конвертація - перша.
MODES = (MODE_CONVERT, MODE_DOCX, MODE_IMAGES, MODE_RENAME, MODE_EXIF)

IMG_EXTS = {".jpg", ".jpeg", ".png", ".gif", ".bmp", ".tiff", ".tif", ".webp"}
# Word (OOXML) не підтримує webp, тому в .docx його не чіпаємо.
DOCX_IMG_EXTS = {".jpg", ".jpeg", ".png", ".gif", ".bmp", ".tif", ".tiff"}
SORT_OPTIONS = ("За іменем", "За датою EXIF", "За датою файлу", "Без сортування")

ORIENTATION_TAG = 0x0112
TMP_SUFFIX = ".tmp_c"
_BAD_NAME_CHARS = re.compile(r'[<>:"/\\|?*\x00-\x1f]')
_CONTENT_TYPES = {".jpg": "image/jpeg", ".jpeg": "image/jpeg", ".png": "image/png"}
_UNSUPPORTED_MODES = {"I", "I;16", "I;16L", "I;16B", "I;16N", "F", "CMYK", "LAB", "HSV", "YCbCr"}
WD_FORMAT_DOCX = 16  # wdFormatDocumentDefault
WD_COMPAT_CURRENT = 65535  # wdCurrent: без режиму сумісності
_EMU_PER_INCH = 914400
_MIN_SIDE = 64  # менше за це картинки в .docx не зменшуємо

# --------------------------------------------------------------------------- #
# Журнал
# --------------------------------------------------------------------------- #

LOG = logging.getLogger("techofficetool")
LOG_PATH: Optional[Path] = None


def setup_logging() -> Optional[Path]:
    """Журнал у %LOCALAPPDATA%\\TechOfficeTool\\TechOfficeTool.log (або у temp)."""
    global LOG_PATH
    if LOG_PATH is not None:
        return LOG_PATH
    try:
        base = Path(os.environ.get("LOCALAPPDATA") or tempfile.gettempdir()) / "TechOfficeTool"
        base.mkdir(parents=True, exist_ok=True)
        path = base / "TechOfficeTool.log"
        handler = logging.handlers.RotatingFileHandler(
            path, maxBytes=512_000, backupCount=2, encoding="utf-8"
        )
        handler.setFormatter(logging.Formatter("%(asctime)s %(levelname)s %(message)s"))
        LOG.addHandler(handler)
        LOG.setLevel(logging.INFO)
        LOG_PATH = path
    except Exception:  # журнал не повинен ламати програму
        LOG.addHandler(logging.NullHandler())
    return LOG_PATH


# --------------------------------------------------------------------------- #
# Дрібні утиліти
# --------------------------------------------------------------------------- #


def is_valid_docx(path: Path) -> bool:
    try:
        with open(path, "rb") as f:
            return f.read(4) == b"PK\x03\x04"
    except Exception:
        return False


def is_valid_docx_deep(path: Path) -> bool:
    """Суворіша перевірка результату конвертації/стиснення."""
    try:
        with zipfile.ZipFile(path) as zf:
            return "word/document.xml" in set(zf.namelist()) and zf.testzip() is None
    except Exception:
        return False


def _norm(p) -> str:
    return os.path.normcase(os.path.abspath(str(p)))


def _same_path(a, b) -> bool:
    return _norm(a) == _norm(b)


def _is_under(path, folder) -> bool:
    try:
        Path(_norm(path)).relative_to(_norm(folder))
        return True
    except ValueError:
        return False


def _tmp_path(dst: Path, suffix: str = TMP_SUFFIX) -> Path:
    return dst.with_name(f".{dst.name}.{uuid.uuid4().hex[:8]}{suffix}")


def _unique_path(p: Path) -> Path:
    if not p.exists():
        return p
    k = 1
    while True:
        cand = p.with_name(f"{p.stem}_{k}{p.suffix}")
        if not cand.exists():
            return cand
        k += 1


class OutputNames:
    """Стежить, щоб два файли одного запуску не записались під одним ім'ям
    (напр. a.bmp і a.jpg - обидва стали б a.jpg).

    primary - "рідні" імена всіх файлів запуску (без зміни розширення);
    чужий файл не може зайняти рідне ім'я іншого.
    """

    def __init__(self, primary=()):
        self.primary = {_norm(p) for p in primary}
        self.used: set = set()

    def claim(self, p: Path, own: bool, avoid_existing: bool) -> Path:
        cand, k = p, 1
        while (
            _norm(cand) in self.used
            or (not own and _norm(cand) in self.primary)
            or (avoid_existing and cand.exists())
        ):
            cand = p.with_name(f"{p.stem}_{k}{p.suffix}")
            k += 1
            own = False  # змінене ім'я вже не "рідне"
        self.used.add(_norm(cand))
        return cand


def _remove_quiet(p: Optional[Path]) -> None:
    if p is None:
        return
    try:
        if p.exists():
            p.unlink()
    except Exception:
        LOG.warning("Не вдалося видалити тимчасовий файл %s", p)


def _copy_times(st: os.stat_result, dst: Path) -> None:
    try:
        os.utime(dst, ns=(st.st_atime_ns, st.st_mtime_ns))
    except Exception:
        pass


def _find_files(
    root: Path,
    recursive: bool,
    exts,
    exclude: Optional[Path] = None,
    cancel=None,
    on_scan: Optional[Callable[[int], None]] = None,
) -> list[Path]:
    """Шукає файли. Пошук можна скасувати (cancel) - на великих дисках він
    може тривати довго, тому подію перевіряємо під час обходу, а не після."""
    it = root.rglob("*") if recursive else root.glob("*")
    out = []
    seen = 0
    for f in it:
        seen += 1
        if seen % 200 == 0:
            if cancel is not None and cancel.is_set():
                break
            if on_scan is not None:
                on_scan(seen)
        try:
            if not f.is_file() or f.name.startswith("~$"):
                continue
        except OSError:
            continue
        if f.suffix.lower() not in exts:
            continue
        if exclude is not None and _is_under(f, exclude):
            continue
        out.append(f)
    out.sort(key=lambda p: str(p).lower())
    return out


# --------------------------------------------------------------------------- #
# Кодування зображень
# --------------------------------------------------------------------------- #


@dataclass
class Encoded:
    data: bytes
    ext: str  # нове розширення з крапкою (нижній регістр)


def _has_alpha(im: Image.Image) -> bool:
    return im.mode in ("RGBA", "LA", "La", "PA", "RGBa") or "transparency" in im.info


def _exif_orientation(im: Image.Image) -> int:
    try:
        v = im.getexif().get(ORIENTATION_TAG, 1)
        return int(v) if v else 1
    except Exception:
        return 1


def _orientation_only_exif(orientation: int) -> Optional[bytes]:
    if not 2 <= int(orientation) <= 8:
        return None
    ex = Image.Exif()
    ex[ORIENTATION_TAG] = int(orientation)
    return ex.tobytes()


def encode_image(
    im: Image.Image,
    ext: str,
    max_side: int,
    quality: int,
    png2jpg: bool,
    exif_mode: str = "keep",
) -> Optional[Encoded]:
    """Повертає нові байти зображення або None, якщо файл краще не чіпати.

    exif_mode:
      "keep"        - зберегти EXIF повністю (орієнтація лишається тегом);
      "orientation" - лишити тільки тег орієнтації (для .docx);
      "strip"       - прибрати EXIF, попередньо повернувши пікселі за орієнтацією.

    Розширення результату завжди відповідає реальному формату: якщо PNG
    став JPEG, ext буде ".jpg".
    """
    ext = ext.lower()
    if getattr(im, "n_frames", 1) > 1 or im.mode in _UNSUPPORTED_MODES:
        return None  # анімація / багатосторінковий TIFF / 16-бітні режими / CMYK

    info = dict(im.info)
    icc = info.get("icc_profile")
    raw_exif = info.get("exif") if isinstance(info.get("exif"), (bytes, bytearray)) else None
    orientation = _exif_orientation(im)

    if exif_mode == "strip" and 2 <= orientation <= 8:
        im = ImageOps.exif_transpose(im)

    if _has_alpha(im):
        alpha = True
        work = im if im.mode == "RGBA" else im.convert("RGBA")
    else:
        alpha = False
        work = im if im.mode in ("RGB", "L") else im.convert("RGB")

    w, h = work.size
    if max(w, h) > max_side:
        ratio = max_side / max(w, h)
        work = work.resize(
            (max(1, int(w * ratio)), max(1, int(h * ratio))), Image.Resampling.LANCZOS
        )

    if ext in (".jpg", ".jpeg"):
        kind, new_ext = ("PNG", ".png") if alpha else ("JPEG", ext)
    elif ext == ".png":
        kind, new_ext = ("PNG", ".png") if (alpha or not png2jpg) else ("JPEG", ".jpg")
    elif ext == ".webp":
        kind, new_ext = "WEBP", ".webp"
    else:  # gif / bmp / tif / tiff
        kind, new_ext = ("PNG", ".png") if alpha else ("JPEG", ".jpg")

    exif_out: Optional[bytes] = None
    if exif_mode == "keep":
        exif_out = bytes(raw_exif) if raw_exif else _orientation_only_exif(orientation)
    elif exif_mode == "orientation":
        exif_out = _orientation_only_exif(orientation)

    buf = io.BytesIO()
    if kind == "JPEG":
        kw = {"quality": quality, "optimize": True, "progressive": True}
    elif kind == "PNG":
        kw = {"optimize": True}
    else:
        kw = {"quality": quality, "method": 4}
    if exif_out:
        kw["exif"] = exif_out
    if icc:
        kw["icc_profile"] = icc
    work.save(buf, kind, **kw)
    return Encoded(buf.getvalue(), new_ext)


def compress_file(
    src: Path,
    dst: Path,
    max_side: int,
    quality: int,
    png2jpg: bool,
    exif_mode: str = "keep",
    names: Optional[OutputNames] = None,
) -> dict:
    """Стискає одне зображення. dst може збігатися з src (стиснення "на місці").

    Файл-результат з'являється атомарно. Якщо результат не менший за оригінал -
    лишається оригінал (виняток: явне видалення EXIF).
    names - спільний облік імен для всього запуску (див. OutputNames).
    """
    res = {"ok": False, "status": "failed", "old": 0, "new": 0, "path": dst}
    in_place = _same_path(src, dst)
    names = names or OutputNames([dst])

    def copy_original():
        # у вихідну папку кладемо незмінену копію (у т.ч. поверх старого результату)
        out = names.claim(dst, own=True, avoid_existing=False)
        out.parent.mkdir(parents=True, exist_ok=True)
        shutil.copy2(src, out)
        res["path"] = out
        return out

    try:
        st = src.stat()
        res["old"] = st.st_size
        ext = src.suffix.lower()
        had_exif = False
        try:
            with Image.open(src) as im:
                im.load()
                had_exif = bool(im.getexif()) or "exif" in im.info
                enc = encode_image(im, ext, max_side, quality, png2jpg, exif_mode)
        except Exception:
            LOG.exception("Не вдалося обробити зображення %s", src)
            if not in_place:
                res["new"] = copy_original().stat().st_size
            else:
                res["new"] = res["old"]
            return res

        use_new = False
        if enc is not None:
            force = exif_mode == "strip" and had_exif
            use_new = force or len(enc.data) < res["old"]

        if not use_new:
            if not in_place:
                copy_original()
            res.update(ok=True, status="skipped" if enc is None else "kept", new=res["old"])
            return res

        same_ext = enc.ext == ext
        if in_place and same_ext:
            final = dst
            names.used.add(_norm(dst))
        else:
            # на місці - не затираємо жодних наявних файлів;
            # у вихідній папці - старий результат попереднього запуску перезаписується
            final = names.claim(
                dst if same_ext else dst.with_suffix(enc.ext), own=same_ext, avoid_existing=in_place
            )
        final.parent.mkdir(parents=True, exist_ok=True)
        tmp = _tmp_path(final)
        try:
            tmp.write_bytes(enc.data)
            _copy_times(st, tmp)
            os.replace(tmp, final)
        finally:
            _remove_quiet(tmp)
        if in_place and not _same_path(final, src):
            try:
                src.unlink()  # PNG -> JPG "на місці": стара копія зайва
            except Exception:
                LOG.warning("Не вдалося видалити замінений оригінал %s", src)
        res.update(ok=True, status="compressed", new=final.stat().st_size, path=final)
    except Exception:
        LOG.exception("Помилка стиснення %s", src)
    return res


def process_one_image(
    src: Path,
    dst: Path,
    max_side: int,
    quality: int,
    png2jpg: bool,
    keep_exif: bool,
    names: Optional[OutputNames] = None,
) -> dict:
    r = compress_file(src, dst, max_side, quality, png2jpg, "keep" if keep_exif else "strip", names)
    return {"ok": r["ok"], "old": r["old"], "new": r["new"], "status": r["status"]}


# --------------------------------------------------------------------------- #
# .docx
# --------------------------------------------------------------------------- #


def _clone_info(info: zipfile.ZipInfo, name: Optional[str] = None) -> zipfile.ZipInfo:
    zi = zipfile.ZipInfo(name or info.filename, date_time=info.date_time)
    zi.compress_type = info.compress_type
    zi.external_attr = info.external_attr
    zi.create_system = info.create_system
    return zi


def _unique_zip_name(name: str, new_ext: str, taken: set) -> str:
    p = PurePosixPath(name)
    cand = str(p.with_suffix(new_ext))
    k = 1
    while cand in taken:
        cand = str(p.with_name(f"{p.stem}_{k}{new_ext}"))
        k += 1
    return cand


def _rewrite_rels(data: bytes, renames: dict) -> bytes:
    for old, new in renames.items():
        o = re.escape(PurePosixPath(old).name.encode("utf-8"))
        n = PurePosixPath(new).name.encode("utf-8")
        pattern = re.compile(rb"""(Target\s*=\s*["'](?:[^"']*/)?media/)""" + o + rb"""(["'])""")
        data = pattern.sub(lambda m, n=n: m.group(1) + n + m.group(2), data)
    return data


def _rewrite_content_types(data: bytes, renames: dict) -> bytes:
    for old, new in renames.items():
        o = re.escape(("/" + old).encode("utf-8"))
        data = re.sub(
            rb'(PartName\s*=\s*")' + o + rb'(")',
            lambda m, new=new: m.group(1) + ("/" + new).encode("utf-8") + m.group(2),
            data,
        )
    exts = {PurePosixPath(n).suffix.lower() for n in renames.values()}
    add = b""
    for ext in sorted(exts):
        ctype = _CONTENT_TYPES.get(ext)
        if not ctype:
            continue
        has = re.search(
            rb'<Default\s[^>]*Extension\s*=\s*"' + re.escape(ext[1:].encode()) + rb'"',
            data,
            re.IGNORECASE,
        )
        if not has:
            add += f'<Default Extension="{ext[1:]}" ContentType="{ctype}"/>'.encode()
    if add:
        data, n = re.subn(rb"(<Types\b[^>]*>)", lambda m: m.group(1) + add, data, count=1)
        if n != 1:
            raise ValueError("[Content_Types].xml має неочікуваний вигляд")
    return data


def _local(tag: str) -> str:
    return tag.rsplit("}", 1)[-1]


def _parse_xml(data: bytes) -> ET.Element:
    if b"<!DOCTYPE" in data[:4096] or b"<!ENTITY" in data[:4096]:
        raise ValueError("XML з DOCTYPE не обробляється")
    return ET.fromstring(data)


def _crop_fraction(value) -> float:
    try:  # a:srcRect: тисячні частки відсотка
        return min(max(int(value), 0), 100000) / 100000
    except (TypeError, ValueError):
        return 0.0


_VML_UNITS = {"in": 1.0, "pt": 1 / 72, "pc": 1 / 6, "cm": 1 / 2.54, "mm": 1 / 25.4, "px": 1 / 96}
_VML_SIZE = re.compile(r"(?:^|;)\s*(width|height)\s*:\s*([0-9.]+)\s*([a-z]*)", re.IGNORECASE)


def _ref_display_size(elem: ET.Element, parents: dict) -> Optional[tuple]:
    """Розмір (ширина, висота) у дюймах, який має покрити ВСЕ зображення для
    цього посилання, або None, якщо його не вдалося надійно визначити."""
    kind = _local(elem.tag)
    if kind == "blip":  # DrawingML: <wp:inline|anchor><wp:extent/> ... <a:blip r:embed/>
        crop_w = crop_h = 1.0
        holder = parents.get(elem)
        if holder is not None:
            for ch in holder:
                name = _local(ch.tag)
                if name == "tile":
                    return None  # заливка-плитка: розмір однієї плитки невідомий
                if name == "srcRect":  # обрізане фото: видно лише частину зображення
                    crop_w = 1.0 - _crop_fraction(ch.get("l")) - _crop_fraction(ch.get("r"))
                    crop_h = 1.0 - _crop_fraction(ch.get("t")) - _crop_fraction(ch.get("b"))
        if crop_w < 0.05 or crop_h < 0.05:
            return None
        node = holder
        while node is not None:
            if _local(node.tag) in ("inline", "anchor"):
                for ch in node:
                    if _local(ch.tag) == "extent":
                        try:
                            cx, cy = int(ch.get("cx")), int(ch.get("cy"))
                        except (TypeError, ValueError):
                            return None
                        if cx > 0 and cy > 0:
                            return (cx / _EMU_PER_INCH / crop_w, cy / _EMU_PER_INCH / crop_h)
                return None
            node = parents.get(node)
        return None
    if kind == "imagedata":  # VML: <v:shape style="width:..;height:.."><v:imagedata r:id/>
        if any(_local(k).lower().startswith("crop") for k in elem.attrib):
            return None
        shape = parents.get(elem)
        if shape is None:
            return None
        dims = {}
        for m in _VML_SIZE.finditer(shape.get("style") or ""):
            unit = _VML_UNITS.get(m.group(3).lower())
            if unit is None:
                return None
            try:
                dims[m.group(1).lower()] = float(m.group(2)) * unit
            except ValueError:
                return None
        if dims.get("width", 0) > 0 and dims.get("height", 0) > 0:
            return (dims["width"], dims["height"])
    return None


def _docx_display_sizes(zin: zipfile.ZipFile, names) -> dict:
    """{ім'я медіафайлу: (ширина, висота) у дюймах} - найбільший розмір, у якому
    зображення показано в документі. Зображення потрапляє сюди, лише якщо
    розмір відомий для КОЖНОГО посилання на нього; інакше діє тільки max_side."""
    sizes: dict = {}
    unknown: set = set()
    for rels_name in sorted(names):
        m = re.match(r"^(.*/)?_rels/([^/]+)\.rels$", rels_name)
        if not m:
            continue
        base = m.group(1) or ""
        rid_map = {}
        for rel in _parse_xml(zin.read(rels_name)):
            if _local(rel.tag) != "Relationship" or (rel.get("TargetMode") or "").lower() == "external":
                continue
            target = rel.get("Target") or ""
            path = target.lstrip("/") if target.startswith("/") else posixpath.normpath(base + target)
            if path.startswith("word/media/"):
                rid_map[rel.get("Id")] = path
        if not rid_map:
            continue
        seen = set()
        try:
            root = _parse_xml(zin.read(base + m.group(2)))
            parents = {child: parent for parent in root.iter() for child in parent}
            for el in root.iter():
                for key, value in el.attrib.items():
                    if value in rid_map and "relationships" in key:
                        seen.add(value)
                        media = rid_map[value]
                        size = _ref_display_size(el, parents)
                        if size is None:
                            unknown.add(media)
                        else:
                            old = sizes.get(media, (0.0, 0.0))
                            sizes[media] = (max(old[0], size[0]), max(old[1], size[1]))
        except Exception:
            LOG.exception("Не вдалося розібрати %s - розміри картинок з нього не враховую", base + m.group(2))
            unknown.update(rid_map.values())
            continue
        unknown.update(media for rid, media in rid_map.items() if rid not in seen)
    return {media: size for media, size in sizes.items() if media not in unknown}


def process_one_docx(
    src: Path, dst: Path, max_side: int, quality: int, png2jpg: bool, ppi: Optional[int] = None
) -> dict:
    """Стискає зображення всередині .docx. Результат - новий zip, зібраний
    запис-за-записом (порядок, дати й стиснення незмінених частин зберігаються).

    ppi - якщо задано, картинка зменшується до цієї щільності відносно того
    розміру, у якому вона реально стоїть на сторінці (але не більше max_side)."""
    result = {"ok": False, "images": 0, "old": 0, "new": 0}
    tmp: Optional[Path] = None
    in_place = _same_path(src, dst)
    try:
        result["old"] = src.stat().st_size
        dst.parent.mkdir(parents=True, exist_ok=True)
        with zipfile.ZipFile(src) as zin:
            infos = zin.infolist()
            names = {i.filename for i in infos}
            replaced: dict[str, bytes] = {}
            renames: dict[str, str] = {}
            taken = set(names)
            display: dict = {}
            if ppi:
                try:
                    display = _docx_display_sizes(zin, names)
                except Exception:
                    LOG.exception("Не вдалося визначити розміри картинок у %s", src)
            for info in infos:
                name = info.filename
                if info.is_dir() or not name.startswith("word/media/"):
                    continue
                ext = PurePosixPath(name).suffix.lower()
                if ext not in DOCX_IMG_EXTS:
                    continue
                try:
                    raw = zin.read(info)
                    with Image.open(io.BytesIO(raw)) as im:
                        im.load()
                        limit = max_side
                        shown = display.get(name)
                        if shown:
                            w, h = im.size
                            # беремо більше зі співвідношень - запас на поворот і неточності
                            scale = max(shown[0] * ppi / w, shown[1] * ppi / h)
                            if scale < 1:
                                limit = min(max_side, max(_MIN_SIDE, math.ceil(max(w, h) * scale)))
                        enc = encode_image(im, ext, limit, quality, png2jpg, "orientation")
                except Exception:
                    LOG.exception("Пропускаю зображення %s у %s", name, src)
                    continue
                if enc is None or len(enc.data) >= len(raw):
                    continue
                new_name = name
                if enc.ext != ext:
                    new_name = _unique_zip_name(name, enc.ext, taken)
                    taken.add(new_name)
                    renames[name] = new_name
                replaced[name] = enc.data

            if not replaced:  # нічого стискати
                if not in_place:
                    shutil.copy2(src, dst)
                result.update(ok=True, new=dst.stat().st_size)
                return result

            if renames and "[Content_Types].xml" not in names:
                raise ValueError("У документі немає [Content_Types].xml")

            tmp = _tmp_path(dst)
            with zipfile.ZipFile(tmp, "w") as zout:
                for info in infos:
                    name = info.filename
                    if name in replaced:
                        zi = _clone_info(info, renames.get(name, name))
                        zi.compress_type = zipfile.ZIP_DEFLATED
                        zout.writestr(zi, replaced[name])
                        continue
                    data = b"" if info.is_dir() else zin.read(info)
                    if renames and name.lower().endswith(".rels"):
                        data = _rewrite_rels(data, renames)
                    elif renames and name == "[Content_Types].xml":
                        data = _rewrite_content_types(data, renames)
                    zout.writestr(_clone_info(info), data)

        # перевірка результату перед підміною
        with zipfile.ZipFile(tmp) as zt:
            if zt.testzip() is not None:
                raise ValueError("Результат пошкоджено (testzip)")
            expected = (names - set(renames)) | set(renames.values())
            if set(zt.namelist()) != expected:
                raise ValueError("Склад архіву після стиснення не збігається з очікуваним")
        if tmp.stat().st_size >= result["old"]:
            _remove_quiet(tmp)
            tmp = None
            if not in_place:
                shutil.copy2(src, dst)
            result.update(ok=True, new=dst.stat().st_size)
            return result
        os.replace(tmp, dst)
        tmp = None
        result["images"] = len(replaced)
        result["new"] = dst.stat().st_size
        result["ok"] = True
    except Exception:
        LOG.exception("Помилка стиснення docx %s", src)
        _remove_quiet(tmp)
        if not in_place:
            try:
                shutil.copy2(src, dst)
                result["new"] = dst.stat().st_size
            except Exception:
                LOG.exception("Не вдалося скопіювати %s у %s", src, dst)
    return result


# --------------------------------------------------------------------------- #
# .doc -> .docx
# --------------------------------------------------------------------------- #


def _delete_original(path: Path) -> None:
    """Видалення оригіналу: у кошик, якщо встановлено send2trash."""
    if _send2trash is not None:
        _send2trash(str(path))
    else:
        os.remove(path)


class WordConverter:
    """Один екземпляр Word на всю пачку файлів (значно швидше, ніж на кожен файл)."""

    def __init__(self, modern: bool = False):
        self.word = None
        self.modern = modern  # True: зберігати без режиму сумісності
        self._com_inited = False

    def __enter__(self):
        try:
            self._start()
        except Exception:
            self.close()
            raise
        return self

    def __exit__(self, *_exc):
        self.close()
        return False

    def _start(self):
        if not self._com_inited:
            pythoncom.CoInitialize()
            self._com_inited = True
        self.word = win32com.client.DispatchEx("Word.Application")
        self.word.Visible = False
        self.word.DisplayAlerts = 0
        try:
            self.word.AutomationSecurity = 3  # msoAutomationSecurityForceDisable: без макросів
        except Exception:
            pass

    def _ensure_alive(self):
        try:
            _ = self.word.Version
        except Exception:
            LOG.warning("Word більше не відповідає - запускаю заново")
            self._quit_quiet()
            self._start()

    def _quit_quiet(self):
        if self.word is not None:
            try:
                self.word.Quit()
            except Exception:
                pass
            self.word = None

    def close(self):
        self._quit_quiet()
        if self._com_inited:
            try:
                pythoncom.CoUninitialize()
            except Exception:
                pass
            self._com_inited = False

    def convert(self, src: Path, dst: Path) -> dict:
        """Конвертує src (.doc) у .docx поруч із dst. Файл з'являється лише після перевірки."""
        result = {"ok": False, "action": "fail", "old": 0, "new": 0, "out": None, "error": "", "macros": False}
        doc = None
        tmp: Optional[Path] = None
        try:
            result["old"] = src.stat().st_size
            self._ensure_alive()
            out = dst.with_suffix(".docx")
            out.parent.mkdir(parents=True, exist_ok=True)
            if out.exists():
                out = _unique_path(out)
            tmp = out.with_name(f".{out.stem}.{uuid.uuid4().hex[:8]}{TMP_SUFFIX}.docx")
            doc = self.word.Documents.Open(
                FileName=str(src.resolve()),
                ConfirmConversions=False,
                ReadOnly=True,
                AddToRecentFiles=False,
                NoEncodingDialog=True,
                PasswordDocument="\u0001no-password\u0001",  # захищений файл -> помилка, а не діалог
                WritePasswordDocument="\u0001no-password\u0001",
            )
            try:  # у .docx макроси не зберігаються - про це треба сказати користувачу
                result["macros"] = bool(doc.HasVBProject)
            except Exception:
                pass
            saved = False
            if self.modern:
                try:
                    doc.SaveAs2(
                        FileName=str(tmp.resolve()),
                        FileFormat=WD_FORMAT_DOCX,
                        CompatibilityMode=WD_COMPAT_CURRENT,
                    )
                    saved = True
                except Exception:
                    LOG.warning("SaveAs2 без режиму сумісності не спрацював для %s - зберігаю звичайним способом", src)
                    _remove_quiet(tmp)
            if not saved:
                doc.SaveAs(FileName=str(tmp.resolve()), FileFormat=WD_FORMAT_DOCX)
            doc.Close(SaveChanges=False)
            doc = None
            if not tmp.exists() or tmp.stat().st_size <= 0 or not is_valid_docx_deep(tmp):
                raise ValueError("Word створив некоректний .docx")
            os.replace(tmp, out)
            tmp = None
            result.update(ok=True, action="convert", new=out.stat().st_size, out=out)
        except Exception as e:
            result["error"] = f"{type(e).__name__}: {e}"
            LOG.exception("Не вдалося конвертувати %s", src)
            try:  # як і раніше: щоб у вихідній папці був хоч якийсь файл
                fallback = dst.parent / src.name
                if not _same_path(fallback, src):
                    fallback.parent.mkdir(parents=True, exist_ok=True)
                    shutil.copy2(src, fallback)
            except Exception:
                LOG.warning("Не вдалося скопіювати %s", src)
        finally:
            if doc is not None:
                try:
                    doc.Close(SaveChanges=False)
                except Exception:
                    pass
            _remove_quiet(tmp)
        return result


# --------------------------------------------------------------------------- #
# Перейменування
# --------------------------------------------------------------------------- #


def _get_exif_datetime(im: Image.Image):
    try:
        dt = im.getexif().get(306)
        if dt:
            return datetime.strptime(dt, "%Y:%m:%d %H:%M:%S")
    except Exception:
        pass
    return None


def _get_date_for_sort(p: Path, mode: str):
    if mode.startswith("За датою EXIF"):
        try:
            with Image.open(p) as im:
                dt = _get_exif_datetime(im)
                if dt:
                    return dt.timestamp()
        except Exception:
            pass
        return p.stat().st_mtime
    if mode.startswith("За датою файлу"):
        return p.stat().st_mtime
    return 0


def _get_exif_date_str(p: Path) -> str:
    try:
        with Image.open(p) as im:
            dt = _get_exif_datetime(im)
            if dt:
                return dt.strftime("%Y-%m-%d")
    except Exception:
        pass
    return datetime.fromtimestamp(p.stat().st_mtime).strftime("%Y-%m-%d")


def validate_template(template: str) -> Optional[str]:
    """None, якщо шаблон коректний, інакше - текст помилки."""
    if not template.strip():
        return "Шаблон порожній."
    try:
        out = template.format(n=1, name="a", parent="p", date="2000-01-01", ext="jpg")
    except (KeyError, IndexError, ValueError, AttributeError) as e:
        return f"Некоректний шаблон ({type(e).__name__}: {e}). Доступно: {{n}}, {{name}}, {{parent}}, {{date}}, {{ext}}."
    if not _BAD_NAME_CHARS.sub("", out).strip(" ."):
        return "Шаблон дає порожню назву."
    return None


def format_name(template: str, n: int, src: Path) -> str:
    try:
        name = template.format(
            n=n,
            name=src.stem,
            parent=src.parent.name,
            date=_get_exif_date_str(src),
            ext=src.suffix.lstrip("."),
        )
    except Exception:
        LOG.warning("Шаблон '%s' не спрацював для %s", template, src)
        name = f"{src.stem}_{n}"
    name = _BAD_NAME_CHARS.sub("_", name).strip(" .")
    return name or f"{src.stem}_{n}"


def plan_renames(files, template: str, start: int, sort_mode: str) -> list:
    files = list(files)
    if sort_mode == "За іменем":
        files.sort(key=lambda p: p.name.lower())
    elif sort_mode.startswith("За датою"):
        files.sort(key=lambda p: _get_date_for_sort(p, sort_mode))
    src_keys = {_norm(f) for f in files}
    used: set = set()
    planned = []
    n = start
    for src in files:
        base = format_name(template, n, src)
        ext = src.suffix
        cand = src.parent / f"{base}{ext}"
        k = 1
        while True:
            key = _norm(cand)
            # файл, який і сам буде перейменовано, не вважається перешкодою
            if key not in used and not (cand.exists() and key not in src_keys):
                break
            cand = src.parent / f"{base}_{k}{ext}"
            k += 1
        used.add(key)
        planned.append((src, cand))
        n += 1
    return planned


def _rename_one(src: Path, dst: Path, attempts: int = 5, delay: float = 0.5) -> None:
    if dst.exists() and _norm(dst) != _norm(src):
        raise FileExistsError(f"{dst.name} вже існує")
    for attempt in range(attempts):
        try:
            os.rename(src, dst)
            return
        except PermissionError:  # файл на мить зайнятий (антивірус, індексатор)
            if attempt == attempts - 1:
                raise
            time.sleep(delay)


def _rename_many(pairs, cancel=None, progress=None):
    """Перейменовує пари (src, dst). Якщо цілі перетинаються з джерелами
    (ланцюжки A->B, B->C або зміна регістру), працює у два етапи через
    тимчасові імена, тому нічого не затирається. Повертає (виконані, помилки)."""
    pairs = [(s, d) for s, d in pairs if str(s) != str(d)]
    done: list = []
    errors: list = []
    total = len(pairs)
    src_keys = {_norm(s) for s, _ in pairs}
    chain = any(_norm(d) in src_keys for _, d in pairs)

    if not chain:
        for i, (s, d) in enumerate(pairs, 1):
            if cancel is not None and cancel.is_set():
                break
            if progress:
                progress(i, total, s.name)
            try:
                _rename_one(s, d)
                done.append((s, d))
            except Exception as e:
                LOG.warning("Не вдалося перейменувати %s: %s", s, e)
                errors.append(f"{s.name}: {e}")
        return done, errors

    staged = []
    try:
        for i, (s, d) in enumerate(pairs, 1):
            if progress:
                progress(i, total, s.name)
            t = s.with_name(f".{s.name}.{uuid.uuid4().hex[:8]}.tot_tmp")
            _rename_one(s, t)
            staged.append((s, t, d))
    except Exception as e:
        LOG.warning("Етап 1 не вдався: %s - відкочую", e)
        errors.append(f"Перейменування скасовано, файли повернуто: {e}")
        for s, t, _ in staged:
            try:
                os.rename(t, s)
            except Exception:
                errors.append(f"Не вдалося повернути {t} -> {s}")
        return [], errors
    for s, t, d in staged:
        try:
            _rename_one(t, d)
            done.append((s, d))
        except Exception as e:
            LOG.warning("Етап 2 не вдався для %s: %s", s, e)
            try:
                os.rename(t, s)
                errors.append(f"{s.name}: {e}")
            except Exception:
                errors.append(f"{s.name}: {e}. Файл лишився під іменем {t.name}")
    return done, errors


# --------------------------------------------------------------------------- #
# Видалення EXIF
# --------------------------------------------------------------------------- #


def _has_metadata(im: Image.Image, keep_icc: bool, paranoid: bool) -> bool:
    info = im.info
    try:
        ex = im.getexif()
        tags = set(ex.keys())
        orient = ex.get(ORIENTATION_TAG, 1)
    except Exception:
        tags, orient = set(), 1
    if tags - {ORIENTATION_TAG}:
        return True
    if paranoid and ORIENTATION_TAG in tags and orient not in (1, None):
        return True
    for k in ("exif", "xmp", "comment", "photoshop"):
        if info.get(k):
            if k == "exif" and not (tags - {ORIENTATION_TAG}) and not paranoid:
                continue  # у "exif" лишився тільки тег орієнтації
            return True
    if any(str(k).startswith("XML:") for k in info):
        return True
    if getattr(im, "text", None):
        return True
    if not keep_icc and info.get("icc_profile"):
        return True
    return False


def strip_jpeg_lossless(data: bytes, keep_icc: bool, orientation: int = 1) -> Optional[bytes]:
    """Прибирає метадані з JPEG без перекодування пікселів. None - якщо файл
    нестандартний (тоді використовується повна переупаковка)."""
    n = len(data)
    if data[:2] != b"\xff\xd8":
        return None
    exif_seg = None
    ox = _orientation_only_exif(orientation) if orientation else None
    if ox:
        payload = ox if ox.startswith(b"Exif\x00\x00") else b"Exif\x00\x00" + ox
        exif_seg = b"\xff\xe1" + struct.pack(">H", len(payload) + 2) + payload
    out = bytearray(b"\xff\xd8")
    inserted = exif_seg is None
    i = 2
    while i < n:
        if data[i] != 0xFF:
            return None
        while i < n and data[i] == 0xFF:
            i += 1
        if i >= n:
            return None
        marker = data[i]
        i += 1
        if marker == 0xD9:  # EOI
            out += b"\xff\xd9"
            return bytes(out) if not data[i:].strip(b"\x00") else None  # хвіст (MPO/HDR) -> повна переупаковка
        if marker == 0x01 or 0xD0 <= marker <= 0xD7:
            out += bytes([0xFF, marker])
            continue
        if i + 2 > n:
            return None
        length = struct.unpack(">H", data[i : i + 2])[0]
        if length < 2 or i + length > n:
            return None
        seg = data[i : i + length]
        i += length
        if marker == 0xDA:  # SOS: далі йдуть закодовані дані до наступного маркера
            if not inserted:
                out += exif_seg
                inserted = True
            out += b"\xff\xda" + seg
            j = i
            while True:
                k = data.find(b"\xff", j)
                if k == -1 or k + 1 >= n:
                    return None
                nxt = data[k + 1]
                if nxt == 0x00 or 0xD0 <= nxt <= 0xD7:
                    j = k + 2
                elif nxt == 0xFF:
                    j = k + 1
                else:
                    break
            out += data[i:k]
            i = k
            continue
        drop = False
        if marker in (0xE1, 0xED, 0xFE, 0xEF) or 0xE3 <= marker <= 0xEC:
            drop = True  # Exif/XMP, IPTC/Photoshop, коментарі, прикладні сегменти
        elif marker == 0xE2:  # ICC або службові (MPF, FlashPix)
            drop = not (keep_icc and seg[2:14] == b"ICC_PROFILE\x00")
        if drop:
            continue
        if marker != 0xE0 and not inserted:
            out += exif_seg
            inserted = True
        out += bytes([0xFF, marker]) + seg
    return None


def _same_pixels(a: bytes, b: bytes) -> bool:
    with Image.open(io.BytesIO(a)) as ia, Image.open(io.BytesIO(b)) as ib:
        ia.load()
        ib.load()
        return ia.mode == ib.mode and ia.size == ib.size and ia.tobytes() == ib.tobytes()


def _repack_clean(im: Image.Image, fmt: str, keep_icc: bool, paranoid: bool) -> bytes:
    icc = im.info.get("icc_profile") if keep_icc else None
    has_alpha = _has_alpha(im)
    base = ImageOps.exif_transpose(im)  # орієнтація зникне разом з EXIF - тож повертаємо пікселі
    if paranoid:
        clean = Image.new("RGBA" if has_alpha else "RGB", base.size)
        clean.paste(base.convert(clean.mode))
    else:
        clean = base.copy()
    for k in list(clean.info):
        if k in ("exif", "xmp", "comment", "photoshop", "icc_profile") or str(k).startswith("XML:"):
            del clean.info[k]
        elif isinstance(clean.info.get(k), str):  # текстові чанки PNG тощо
            del clean.info[k]
    if fmt == "MPO":
        fmt = "JPEG"
    kw = {}
    if fmt in ("JPEG", "JPG"):
        fmt, kw = "JPEG", {"quality": 95, "optimize": True}
    elif fmt == "PNG":
        kw = {"optimize": True, "compress_level": 9}
    elif fmt == "WEBP":
        kw = {"quality": 95, "method": 6}
    if icc and fmt in ("JPEG", "PNG", "WEBP"):
        kw["icc_profile"] = icc
    buf = io.BytesIO()
    clean.save(buf, fmt, **kw)
    return buf.getvalue()


def strip_exif_one(p: Path, keep_icc: bool, paranoid: bool, backup: bool) -> str:
    """'cleaned' | 'already_clean' | 'skipped' | 'error'. Оригінал підміняється
    атомарно, лише після перевірки, що метаданих у результаті справді немає."""
    tmp: Optional[Path] = None
    try:
        st = p.stat()
        raw = p.read_bytes()
        with Image.open(io.BytesIO(raw)) as im:
            fmt = im.format or "JPEG"
            # GIF/BMP/TIFF: EXIF там або відсутній, або змішаний зі службовими тегами,
            # тому безпечно очистити їх не можна. Анімацію переупаковка зруйнувала б.
            if fmt not in ("JPEG", "MPO", "PNG", "WEBP"):
                return "skipped"
            if fmt != "MPO" and getattr(im, "n_frames", 1) > 1:
                return "skipped"
            if not _has_metadata(im, keep_icc, paranoid):
                return "already_clean"
            orientation = _exif_orientation(im)
            im.load()
            new = None
            if fmt in ("JPEG", "MPO") and not paranoid:
                try:
                    cand = strip_jpeg_lossless(raw, keep_icc, orientation)
                    if cand is not None and _same_pixels(raw, cand):
                        new = cand
                except Exception:
                    LOG.exception("Безвтратне очищення не вдалося для %s", p)
            if new is None:
                new = _repack_clean(im, fmt, keep_icc, paranoid)

        with Image.open(io.BytesIO(new)) as chk:  # результат має бути чистим
            if _has_metadata(chk, keep_icc, paranoid):
                raise ValueError("метадані лишилися після очищення")
        if backup:
            bak = p.with_suffix(p.suffix + ".bak")
            if not bak.exists():
                shutil.copy2(p, bak)
        tmp = _tmp_path(p)
        tmp.write_bytes(new)
        _copy_times(st, tmp)
        os.replace(tmp, p)
        tmp = None
        return "cleaned"
    except Exception:
        LOG.exception("Не вдалося очистити EXIF у %s", p)
        return "error"
    finally:
        _remove_quiet(tmp)


# --------------------------------------------------------------------------- #
# Запуск завдання (використовується GUI)
# --------------------------------------------------------------------------- #

ProgressFn = Callable[[int, int, str], None]


@dataclass
class JobConfig:
    mode: str
    input_dir: Path
    output_dir: Path
    level: Optional[dict] = None
    recursive: bool = False
    keep_exif: bool = True
    keep_icc: bool = True
    paranoid: bool = False
    backup: bool = True
    dry_run: bool = True
    delete_doc: bool = False
    modern_format: bool = False
    template: str = "IMG_{n:04d}"
    start_num: int = 1
    sort: str = SORT_OPTIONS[0]


def _mb(x: int) -> str:
    return f"{x / 1024 / 1024:.1f} MB"


def _cancelled(cancel) -> bool:
    return bool(cancel is not None and cancel.is_set())


def run_job(
    cfg: JobConfig,
    progress: Optional[ProgressFn] = None,
    cancel=None,
    converter_factory: Callable[[], WordConverter] = WordConverter,
) -> str:
    """Виконує завдання й повертає підсумкове повідомлення для користувача."""
    progress = progress or (lambda i, t, n: None)
    src_root, out_root = Path(cfg.input_dir), Path(cfg.output_dir)
    exclude = None if _same_path(src_root, out_root) else out_root
    suffix = "\n\nСкасовано користувачем - оброблено {i} з {t}."
    cancelled_scan = "Скасовано користувачем під час пошуку файлів."

    def find(recursive, exts, excl=None):
        return _find_files(
            src_root,
            recursive,
            exts,
            excl,
            cancel,
            lambda n: progress(0, 0, f"Пошук файлів… переглянуто {n}"),
        )

    if cfg.mode == MODE_DOCX:
        files = [
            f
            for f in find(True, {".docx"}, exclude)
            if not f.name.endswith(".tmp_c.docx") and is_valid_docx(f)
        ]
        if _cancelled(cancel):
            return cancelled_scan
        if not files:
            return "Не знайдено .docx"
        lv = cfg.level
        total = len(files)
        old_sum = new_sum = img_sum = errs = done = 0
        for i, src in enumerate(files, 1):
            if _cancelled(cancel):
                break
            progress(i, total, src.name[:50])
            r = process_one_docx(
                src,
                out_root / src.relative_to(src_root),
                lv["max_side"],
                lv["quality"],
                lv["png2jpg"],
                lv.get("ppi"),
            )
            done += 1
            old_sum += r["old"]
            new_sum += r["new"]
            img_sum += r["images"]
            if not r["ok"]:
                errs += 1
        msg = f"Готово! Файлів: {done}, Зображень: {img_sum}\nБуло: {_mb(old_sum)}, Стало: {_mb(new_sum)}"
        if old_sum > 0:
            msg += f"\nЗменшення: {(1 - new_sum / old_sum) * 100:.1f}%"
        if errs:
            msg += f"\nПомилок: {errs}"
        if done < total:
            msg += suffix.format(i=done, t=total)
        return msg

    if cfg.mode == MODE_CONVERT:
        if converter_factory is WordConverter and not WIN32_OK:
            return "Для конвертації потрібен pywin32."
        files = find(cfg.recursive, {".doc"}, exclude)
        if _cancelled(cancel):
            return cancelled_scan
        if not files:
            return "У папці не знайдено файлів .doc для конвертації."
        total = len(files)
        ok = fail = deleted = done = exists = 0
        err_files: list[str] = []
        macro_files: list[str] = []
        try:
            ctx = converter_factory()
            conv = ctx.__enter__()
            conv.modern = cfg.modern_format
        except Exception as e:
            LOG.exception("Не вдалося запустити Word")
            return (
                "Не вдалося запустити Microsoft Word.\n"
                f"Перевірте, що Word встановлено й активовано.\n\n({type(e).__name__}: {e})"
            )
        try:
            for i, src in enumerate(files, 1):
                if _cancelled(cancel):
                    break
                progress(i, total, src.name[:50])
                dst = out_root / src.relative_to(src_root)
                done += 1
                if dst.with_suffix(".docx").exists():
                    exists += 1  # вже конвертовано раніше - не дублюємо і не перезаписуємо
                    continue
                r = conv.convert(src, dst)
                if r["ok"]:
                    ok += 1
                    if r.get("macros"):
                        macro_files.append(src.name)  # оригінал з макросами не видаляємо
                    elif cfg.delete_doc and r["out"] is not None and r["new"] > 10:
                        try:
                            if is_valid_docx_deep(r["out"]):
                                _delete_original(src)
                                deleted += 1
                        except Exception:
                            LOG.exception("Не вдалося видалити оригінал %s", src)
                else:
                    fail += 1
                    err_files.append(f"{src.name} ({r['error'][:80]})" if r["error"] else src.name)
        finally:
            ctx.__exit__(None, None, None)
        msg = f"Готово! Конвертовано: {ok}, Помилок: {fail}"
        if exists:
            msg += f"\nПропущено (поруч уже є .docx): {exists}"
        if cfg.delete_doc and deleted > 0:
            where = "у кошик" if _send2trash is not None else "назавжди"
            msg += f"\nВидалено оригіналів .doc ({where}): {deleted}"
        if macro_files:
            msg += (
                f"\n\nМістили макроси: {len(macro_files)}. У .docx макросів немає,"
                " тому оригінали .doc цих файлів збережено:\n" + "\n".join(macro_files[:5])
            )
        if fail > 0:
            msg += "\n\nНе вдалося конвертувати:\n" + "\n".join(err_files[:5])
        if done < total:
            msg += suffix.format(i=done, t=total)
        return msg

    if cfg.mode == MODE_IMAGES:
        files = find(cfg.recursive, IMG_EXTS, exclude)
        if _cancelled(cancel):
            return cancelled_scan
        if not files:
            return "Немає зображень"
        lv = cfg.level
        total = len(files)
        old_sum = new_sum = ok = err = skipped = done = 0
        err_files = []
        names = OutputNames(out_root / f.relative_to(src_root) for f in files)
        for i, src in enumerate(files, 1):
            if _cancelled(cancel):
                break
            progress(i, total, src.name[:50])
            r = process_one_image(
                src,
                out_root / src.relative_to(src_root),
                lv["max_side"],
                lv["quality"],
                lv["png2jpg"],
                cfg.keep_exif,
                names,
            )
            done += 1
            if r["ok"]:
                ok += 1
                old_sum += r["old"]
                new_sum += r["new"]
                if r["status"] == "skipped":
                    skipped += 1
            else:
                err += 1
                err_files.append(src.name)
        msg = f"Готово! Оброблено: {ok}/{total}\nБуло: {_mb(old_sum)}, Стало: {_mb(new_sum)}"
        if old_sum > 0:
            msg += f"\nЗменшення: {(1 - new_sum / old_sum) * 100:.1f}%"
        if skipped:
            msg += f"\nБез змін (анімація, CMYK або 16-бітні): {skipped}"
        if err:
            msg += f"\nПомилок: {err}\nНе вдалося обробити:\n" + "\n".join(err_files[:5])
        if done < total:
            msg += suffix.format(i=done, t=total)
        return msg

    if cfg.mode == MODE_RENAME:
        tmpl = cfg.template.strip() or "IMG_{n:04d}"
        bad = validate_template(tmpl)
        if bad:
            return bad
        files = find(cfg.recursive, IMG_EXTS)
        if _cancelled(cancel):
            return cancelled_scan
        if not files:
            return "У вибраній папці не знайдено зображень для перейменування."
        planned = plan_renames(files, tmpl, cfg.start_num, cfg.sort)
        if cfg.dry_run:
            lines = [f"Знайдено: {len(planned)}", "План (перші 50):"]
            for s, d in planned[:50]:
                lines.append(f"{s.name} → {d.name}")
            if len(planned) > 50:
                lines.append(f"... ще {len(planned) - 50}")
            return "\n".join(lines)
        done_pairs, errors = _rename_many(planned, cancel, progress)
        msg = f"Перейменовано: {len(done_pairs)}\nПомилок: {len(errors)}"
        if errors:
            msg += "\n\nДеталі помилок:\n" + "\n".join(errors[:3])
        return msg

    if cfg.mode == MODE_EXIF:
        files = find(cfg.recursive, IMG_EXTS)
        if _cancelled(cancel):
            return cancelled_scan
        if not files:
            return "Немає зображень"
        total = len(files)
        cleaned = already = errs = done = skipped = 0
        err_files = []
        for i, src in enumerate(files, 1):
            if _cancelled(cancel):
                break
            progress(i, total, src.name[:50])
            r = strip_exif_one(src, cfg.keep_icc, cfg.paranoid, cfg.backup)
            done += 1
            if r == "cleaned":
                cleaned += 1
            elif r == "already_clean":
                already += 1
            elif r == "skipped":
                skipped += 1
            else:
                errs += 1
                err_files.append(src.name)
        msg = f"Очищено: {cleaned}\nБез EXIF: {already}\nПомилок: {errs}\nВсього: {total}"
        if skipped:
            msg += f"\nПропущено (GIF/BMP/TIFF або анімація): {skipped}"
        if errs:
            msg += "\n\nНе вдалося обробити:\n" + "\n".join(err_files[:5])
        if cfg.backup:
            msg += "\nОригінали збережено як *.bak (УВАГА: у них EXIF/GPS лишається!)"
        if done < total:
            msg += suffix.format(i=done, t=total)
        return msg

    return f"Невідомий режим: {cfg.mode}"

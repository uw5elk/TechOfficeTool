import io
import pytest
from PIL import Image


@pytest.fixture
def noise():
    """Генератор важких для стиснення зображень (щоб було що зменшувати)."""

    def make(size=(1800, 1200), mode="RGB"):
        return Image.effect_noise(size, 64).convert(mode)

    return make


def jpeg_bytes(img, **kw):
    buf = io.BytesIO()
    img.save(buf, "JPEG", **kw)
    return buf.getvalue()

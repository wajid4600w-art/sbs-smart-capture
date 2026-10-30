import io

import pytest
from PIL import Image

from app import MAX_GEMINI_IMAGE_BYTES, normalize_image_for_gemini


def image_bytes(image_format):
    image = Image.new("RGB", (40, 30), (40, 120, 200))
    output = io.BytesIO()
    image.save(output, format=image_format)
    return output.getvalue()


@pytest.mark.parametrize("image_format", ["PNG", "JPEG", "WEBP"])
def test_supported_images_are_normalized_with_matching_mime(image_format):
    normalized, mime = normalize_image_for_gemini(image_bytes(image_format))

    assert mime == "image/jpeg"
    assert len(normalized) <= MAX_GEMINI_IMAGE_BYTES
    with Image.open(io.BytesIO(normalized)) as image:
        assert image.format == "JPEG"


def test_corrupt_image_is_rejected():
    with pytest.raises(ValueError, match="corrupt or unsupported"):
        normalize_image_for_gemini(b"not an image")


def test_empty_image_is_rejected():
    with pytest.raises(ValueError, match="empty"):
        normalize_image_for_gemini(b"")


@pytest.mark.parametrize(
    ("source_format", "clipboard_mime"),
    [("PNG", "image/jpeg"), ("JPEG", "image/webp"), ("WEBP", "image/png")],
)
def test_pasted_images_use_bytes_not_clipboard_mime(source_format, clipboard_mime):
    # Browser clipboard MIME metadata can be wrong; normalization inspects the bytes.
    normalized, mime = normalize_image_for_gemini(image_bytes(source_format))

    assert clipboard_mime.startswith("image/")
    assert mime == "image/jpeg"
    with Image.open(io.BytesIO(normalized)) as image:
        assert image.format == "JPEG"

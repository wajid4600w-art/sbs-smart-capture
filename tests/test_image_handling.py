import io

import pytest
from PIL import Image

from app import MAX_GEMINI_IMAGE_BYTES, MAX_IMAGE_PIXELS, normalize_image_for_gemini


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


def test_oversized_pixel_dimensions_are_rejected_before_decode():
    image = Image.new("RGB", (5001, 5001), (40, 120, 200))
    output = io.BytesIO()
    image.save(output, format="JPEG", quality=1)

    with pytest.raises(ValueError, match="too many pixels"):
        normalize_image_for_gemini(output.getvalue())


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


def test_gemini_receives_already_normalized_bytes_without_second_normalization(monkeypatch):
    import app

    normalized, _ = normalize_image_for_gemini(image_bytes("PNG"))
    calls = []

    class FakeModels:
        def generate_content(self, **kwargs):
            calls.append(kwargs["contents"][1])
            return type(
                "Response",
                (),
                {
                    "text": (
                        '{"record_type":"General Document","name":"Test","shop":"N/A",'
                        '"doc_date":"2026-10-30","items":"N/A","total_amount":"N/A",'
                        '"plain_text":"Test"}'
                    )
                },
            )()

    class FakeClient:
        models = FakeModels()

    monkeypatch.setattr(app.genai, "Client", lambda api_key: FakeClient())
    monkeypatch.setattr(app, "normalize_image_for_gemini", lambda _: (_ for _ in ()).throw(
        AssertionError("image was normalized twice")
    ))
    monkeypatch.setenv("GEMINI_API_KEY", "test-key")

    result = app.analyze_with_gemini(image_bytes=normalized)

    assert result["name"] == "Test"
    assert len(calls) == 1

import io
from types import SimpleNamespace

from PIL import Image

import app_effi


def test_load_invoice_uploaded_file_uses_custom_vision_model_and_timeout(monkeypatch):
    buf = io.BytesIO()
    Image.new("RGB", (20, 20), color="white").save(buf, format="PNG")
    uploaded = SimpleNamespace(name="factura.png", getvalue=lambda: buf.getvalue())

    seen = {}

    def fake_try_vision_extract(images, source_name, model=None, manual_rules=None, timeout=180):
        seen["model"] = model
        seen["timeout"] = timeout
        return [{"description": "Producto", "quantity": 1, "unit_price_invoice": 5, "total_invoice": 5}], "ok"

    monkeypatch.setattr(app_effi, "try_vision_extract", fake_try_vision_extract)

    result = app_effi.load_invoice_uploaded_file(
        uploaded,
        ollama_settings={"model": "text-model"},
        read_mode="Auto (Vision si hay Ollama)",
        manual_rules={},
        vision_model_name="vision-mini",
        vision_timeout=42,
    )

    assert result["kind"] == "vision-ollama"
    assert seen["model"] == "vision-mini"
    assert seen["timeout"] == 42

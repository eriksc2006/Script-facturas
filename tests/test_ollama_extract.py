import io
import json
from urllib.error import HTTPError, URLError

from effi_processor.ollama_extract import (
    MAX_TEXT_CHARS,
    _request_json,
    check_ollama,
    coerce_bool,
    extract_invoice_items,
    extract_json_object,
    items_from_payload,
    max_prompt_chars,
    normalize_item,
    parse_number_loose,
    select_extracted_rows,
    should_call_ollama,
    validate_ollama_base,
)


def test_max_prompt_chars_shrinks_for_moondream():
    assert max_prompt_chars("moondream:latest") == 2800
    assert max_prompt_chars("llama3.2") == MAX_TEXT_CHARS


def test_validate_host_requires_http():
    assert validate_ollama_base("http://127.0.0.1:11434/") == "http://127.0.0.1:11434"
    try:
        validate_ollama_base("file:///tmp")
        assert False, "expected ValueError"
    except ValueError:
        pass


def test_parse_number_loose_accepts_colombian_format():
    assert parse_number_loose("1.234,50") == 1234.50
    assert parse_number_loose("12,5") == 12.5
    assert parse_number_loose(10) == 10.0
    assert parse_number_loose(None) is None
    assert parse_number_loose(True) is None


def test_coerce_bool_spanish():
    assert coerce_bool("sí") is True
    assert coerce_bool("no") is False
    assert coerce_bool(True) is True


def test_extract_json_object_from_fences():
    raw = '```json\n{"items": [{"description": "Agua"}]}\n```'
    data = extract_json_object(raw)
    assert data["items"][0]["description"] == "Agua"


def test_normalize_item_requires_description_or_code():
    assert normalize_item({}) is None
    row = normalize_item(
        {
            "description": "ACEITE 1L",
            "quantity": "2",
            "unit_price_invoice": "1000",
            "is_bonus": "si",
        },
        source_name="f.pdf",
    )
    assert row is not None
    assert row["quantity"] == 2.0
    assert row["is_bonus"] is True
    assert row["extraction_source"] == "ollama"
    assert row["archivo_origen"] == "f.pdf"


def test_select_extracted_rows_fallback_and_always():
    regex = [{"description": "regex"}]
    llm = [{"description": "ollama"}]
    assert select_extracted_rows(regex, llm, "fallback") == regex
    assert select_extracted_rows([], llm, "fallback") == llm
    assert select_extracted_rows(regex, llm, "always") == llm
    assert select_extracted_rows(regex, [], "always") == regex


def test_should_call_ollama_skips_fallback_when_regex_works():
    rows = [{"description": "regex"}]
    assert should_call_ollama(True, "fallback", rows, "texto") is False
    assert should_call_ollama(True, "fallback", [], "texto") is True
    assert should_call_ollama(True, "always", rows, "texto") is True
    assert should_call_ollama(False, "always", [], "texto") is False
    assert should_call_ollama(True, "always", [], "  ") is False


def test_items_from_payload_spanish_alias():
    rows = items_from_payload({"lineas": [{"descripcion": "Jabón", "cantidad": 1}]})
    assert len(rows) == 1
    assert rows[0]["description"] == "Jabón"


def test_check_ollama_ok(monkeypatch):
    def fake_request(url, payload=None, timeout=5.0):
        assert "api/tags" in url
        return {"models": [{"name": "llama3.2:latest"}]}

    monkeypatch.setattr("effi_processor.ollama_extract._request_json", fake_request)
    status = check_ollama("http://127.0.0.1:11434")
    assert status["ok"] is True
    assert "llama3.2:latest" in status["models"]


def test_check_ollama_down(monkeypatch):
    def fake_request(url, payload=None, timeout=5.0):
        raise URLError("connection refused")

    monkeypatch.setattr("effi_processor.ollama_extract._request_json", fake_request)
    status = check_ollama("http://127.0.0.1:11434")
    assert status["ok"] is False
    assert "ollama.com" in status["message"]


def test_extract_invoice_items_success(monkeypatch):
    def fake_request(url, payload=None, timeout=5.0):
        assert payload["model"] == "llama3.2"
        assert payload["format"] == "json"
        user_content = payload["messages"][1]["content"]
        assert "Leche" in user_content
        return {
            "message": {
                "content": json.dumps(
                    {
                        "tax_included": True,
                        "items": [
                            {
                                "code": "7701",
                                "description": "Leche",
                                "quantity": 3,
                                "unit_price_invoice": 2000,
                                "total_invoice": 6000,
                            }
                        ],
                    }
                )
            }
        }

    monkeypatch.setattr("effi_processor.ollama_extract._request_json", fake_request)
    result = extract_invoice_items("FACTURA\nLeche 3 2000 6000", model="llama3.2")
    assert result["ok"] is True
    assert result["tax_included"] is True
    assert result["rows"][0]["description"] == "Leche"
    assert result["rows"][0]["code"] == "7701"


def test_extract_invoice_items_empty_text():
    result = extract_invoice_items("   ")
    assert result["ok"] is False
    assert result["rows"] == []


def test_extract_invoice_items_model_missing(monkeypatch):
    def fake_request(url, payload=None, timeout=5.0):
        raise HTTPError(
            url,
            404,
            "Not Found",
            hdrs=None,
            fp=io.BytesIO(b'{"error":"not found"}'),
        )

    monkeypatch.setattr("effi_processor.ollama_extract._request_json", fake_request)
    result = extract_invoice_items("producto 1", model="llama3.2")
    assert result["ok"] is False
    assert "ollama pull" in result["message"]


def test_extract_invoice_items_connection_error(monkeypatch):
    def fake_request(url, payload=None, timeout=5.0):
        raise URLError("connection refused")

    monkeypatch.setattr("effi_processor.ollama_extract._request_json", fake_request)
    result = extract_invoice_items("producto 1")
    assert result["ok"] is False
    assert "conexión" in result["message"]


def test_extract_invoice_items_timeout(monkeypatch):
    def fake_request(url, payload=None, timeout=5.0):
        raise TimeoutError()

    monkeypatch.setattr("effi_processor.ollama_extract._request_json", fake_request)
    result = extract_invoice_items("producto 1")
    assert result["ok"] is False
    assert "tardó" in result["message"]


def test_extract_invoice_items_invalid_host():
    result = extract_invoice_items("producto", host="not-a-url")
    assert result["ok"] is False


def test_extract_invoice_items_legacy_response_and_tax_text(monkeypatch):
    def fake_request(url, payload=None, timeout=5.0):
        return {
            "response": json.dumps(
                {
                    "tax_included": "sí",
                    "items": [{"description": "Arroz", "quantity": 1}],
                }
            )
        }

    monkeypatch.setattr("effi_processor.ollama_extract._request_json", fake_request)
    result = extract_invoice_items("Arroz")
    assert result["ok"] is True
    assert result["tax_included"] is True


def test_extract_invoice_items_no_lines(monkeypatch):
    def fake_request(url, payload=None, timeout=5.0):
        return {"message": {"content": '{"items": []}'}}

    monkeypatch.setattr("effi_processor.ollama_extract._request_json", fake_request)
    result = extract_invoice_items("totales")
    assert result["ok"] is False
    assert result["rows"] == []


def test_check_ollama_empty_models(monkeypatch):
    monkeypatch.setattr(
        "effi_processor.ollama_extract._request_json",
        lambda url, payload=None, timeout=5.0: {"models": []},
    )
    status = check_ollama("http://127.0.0.1:11434")
    assert status["ok"] is False
    assert "no hay modelos" in status["message"]


def test_check_ollama_http_error(monkeypatch):
    def fake_request(url, payload=None, timeout=5.0):
        raise HTTPError(url, 500, "err", hdrs=None, fp=io.BytesIO(b""))

    monkeypatch.setattr("effi_processor.ollama_extract._request_json", fake_request)
    status = check_ollama("http://127.0.0.1:11434")
    assert status["ok"] is False
    assert "HTTP 500" in status["message"]


def test_check_ollama_invalid_host():
    status = check_ollama("ftp://x")
    assert status["ok"] is False


class _FakeResponse:
    def __init__(self, payload: bytes):
        self._payload = payload

    def read(self):
        return self._payload

    def __enter__(self):
        return self

    def __exit__(self, *args):
        return False


def test_request_json_get_and_post(monkeypatch):
    captured = {}

    def fake_urlopen(request, timeout=0):
        captured["method"] = request.get_method()
        captured["timeout"] = timeout
        return _FakeResponse(b'{"ok": true}')

    monkeypatch.setattr("effi_processor.ollama_extract.urlopen", fake_urlopen)
    body = _request_json("http://127.0.0.1:11434/api/tags", timeout=2)
    assert body == {"ok": True}
    posted = _request_json(
        "http://127.0.0.1:11434/api/chat",
        payload={"model": "llama3.2"},
        timeout=9,
    )
    assert posted == {"ok": True}
    assert captured["method"] == "POST"


def test_extract_json_object_garbage():
    assert extract_json_object("no json here") == {}
    assert extract_json_object("") == {}
    wrapped = 'prefix {"items": []} suffix'
    assert extract_json_object(wrapped)["items"] == []

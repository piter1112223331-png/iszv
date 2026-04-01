import json

from parser.main import _resolve_flowis_url, _select_payload, main


class _Doc:
    def __init__(self):
        self.llm_payload = {"kind": "llm"}
        self.flowis_payload = {"kind": "flowis"}

    def to_dict(self):
        return {"kind": "full"}


def test_select_payload_modes():
    doc = _Doc()
    assert _select_payload(doc, "llm") == {"kind": "llm"}
    assert _select_payload(doc, "flowis") == {"kind": "flowis"}
    assert _select_payload(doc, "full") == {"kind": "full"}


def test_resolve_flowis_url_from_env(monkeypatch):
    monkeypatch.setenv("FLOWIS_API_URL", "http://env-url")
    assert _resolve_flowis_url(None) == "http://env-url"
    assert _resolve_flowis_url("http://cli-url") == "http://cli-url"


def test_missing_flowis_url_returns_error(monkeypatch, tmp_path):
    monkeypatch.setattr("parser.main.parse_notice", lambda path, debug=False: _Doc())
    monkeypatch.delenv("FLOWIS_API_URL", raising=False)
    monkeypatch.setattr(
        "sys.argv",
        ["prog", "dummy.xlsx", "--send-flowis", "--output", str(tmp_path / "out.json")],
    )
    rc = main()
    assert rc == 2


def test_flowis_save_response_logic(monkeypatch, tmp_path):
    monkeypatch.setattr("parser.main.parse_notice", lambda path, debug=False: _Doc())
    monkeypatch.setattr("parser.main.send_to_flowis", lambda *args, **kwargs: {"answer": "ok"})
    out_path = tmp_path / "parsed.json"
    resp_path = tmp_path / "flowis_response.json"
    monkeypatch.setattr(
        "sys.argv",
        [
            "prog",
            "dummy.xlsx",
            "--output",
            str(out_path),
            "--send-flowis",
            "--flowis-url",
            "http://localhost/api",
            "--flowis-save-response",
            str(resp_path),
        ],
    )
    rc = main()
    assert rc == 0
    assert resp_path.exists()
    assert json.loads(resp_path.read_text(encoding="utf-8")) == {"answer": "ok"}

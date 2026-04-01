import json

import pytest

from parser.flowis_client import build_flowis_request_body, send_to_flowis


def test_build_flowis_request_body_question_json():
    payload = {"a": 1, "b": "тест"}
    body = build_flowis_request_body(payload)
    assert "question" in body
    assert json.loads(body["question"]) == payload


def test_send_to_flowis_success(monkeypatch):
    class _Resp:
        def raise_for_status(self):
            return None

        def json(self):
            return {"ok": True}

    def _post(url, json=None, timeout=None):  # noqa: A002
        return _Resp()

    monkeypatch.setattr("parser.flowis_client.requests.post", _post)
    got = send_to_flowis("http://localhost", {"question": "{}"}, timeout_sec=1)
    assert got == {"ok": True}


def test_send_to_flowis_invalid_json(monkeypatch):
    class _Resp:
        def raise_for_status(self):
            return None

        def json(self):
            raise ValueError("bad json")

    def _post(url, json=None, timeout=None):  # noqa: A002
        return _Resp()

    monkeypatch.setattr("parser.flowis_client.requests.post", _post)
    with pytest.raises(RuntimeError, match="not valid JSON"):
        send_to_flowis("http://localhost", {"question": "{}"}, timeout_sec=1)

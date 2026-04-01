from __future__ import annotations

import json

try:
    import requests  # type: ignore
except ModuleNotFoundError:  # pragma: no cover - environment fallback
    class _RequestsFallback:
        class Timeout(Exception):
            pass

        class ConnectionError(Exception):
            pass

        class HTTPError(Exception):
            def __init__(self, response=None):
                super().__init__("HTTP error")
                self.response = response

        @staticmethod
        def post(*args, **kwargs):
            raise RuntimeError("requests dependency is not installed")

    requests = _RequestsFallback()  # type: ignore


def build_flowis_request_body(selected_payload: dict, mode: str = "question_json") -> dict:
    if mode == "question_json":
        return {"question": json.dumps(selected_payload, ensure_ascii=False)}
    raise ValueError(f"Unsupported Flowis request body mode: {mode}")


def send_to_flowis(api_url: str, payload: dict, timeout_sec: int = 180) -> dict:
    try:
        response = requests.post(api_url, json=payload, timeout=timeout_sec)
        response.raise_for_status()
    except requests.Timeout as exc:
        raise RuntimeError(f"Flowis request timed out after {timeout_sec} seconds") from exc
    except requests.ConnectionError as exc:
        raise RuntimeError(f"Flowis connection error for URL: {api_url}") from exc
    except requests.HTTPError as exc:
        status = exc.response.status_code if exc.response is not None else "unknown"
        body = exc.response.text if exc.response is not None else ""
        raise RuntimeError(f"Flowis HTTP error {status}: {body}") from exc

    try:
        return response.json()
    except ValueError as exc:
        raise RuntimeError("Flowis response is not valid JSON") from exc

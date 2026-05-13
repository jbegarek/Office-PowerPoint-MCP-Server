import time
import sys
from types import SimpleNamespace
from urllib.parse import parse_qs, urlparse

import pytest
from starlette.applications import Starlette
from starlette.responses import JSONResponse
from starlette.routing import Route
from starlette.testclient import TestClient

import ppt_mcp_server
import utils.http_auth as http_auth
from utils.http_auth import APIKeyMiddleware, build_download_signature


async def _ok_endpoint(request):
    return JSONResponse({"ok": True})


def _create_app() -> Starlette:
    app = Starlette(
        routes=[
            Route("/protected", _ok_endpoint, methods=["GET", "OPTIONS"]),
            Route("/health", _ok_endpoint, methods=["GET", "OPTIONS"]),
            Route("/healthz", _ok_endpoint, methods=["GET", "OPTIONS"]),
            Route("/files/{filename}", _ok_endpoint, methods=["GET", "OPTIONS"]),
        ]
    )
    app.add_middleware(
        APIKeyMiddleware,
        api_key="expected-secret",
        header_name="x-api-key",
        exempt_paths=["/health", "/healthz"],
    )
    return app


def test_missing_api_key_returns_401(monkeypatch):
    monkeypatch.delenv("DOC_DOWNLOAD_SIGNING_SECRET", raising=False)
    monkeypatch.delenv("PPTX_MCP_API_KEY", raising=False)
    client = TestClient(_create_app())

    response = client.get("/protected")

    assert response.status_code == 401
    assert response.json() == {"error": "Missing API key header: x-api-key"}


def test_invalid_api_key_returns_403(monkeypatch):
    monkeypatch.delenv("DOC_DOWNLOAD_SIGNING_SECRET", raising=False)
    monkeypatch.delenv("PPTX_MCP_API_KEY", raising=False)
    client = TestClient(_create_app())

    response = client.get("/protected", headers={"x-api-key": "wrong"})

    assert response.status_code == 403
    assert response.json() == {"error": "Invalid API key."}


def test_valid_api_key_passes(monkeypatch):
    monkeypatch.delenv("DOC_DOWNLOAD_SIGNING_SECRET", raising=False)
    monkeypatch.delenv("PPTX_MCP_API_KEY", raising=False)
    client = TestClient(_create_app())

    response = client.get("/protected", headers={"x-api-key": "expected-secret"})

    assert response.status_code == 200
    assert response.json() == {"ok": True}


def test_health_paths_are_exempt(monkeypatch):
    monkeypatch.delenv("DOC_DOWNLOAD_SIGNING_SECRET", raising=False)
    monkeypatch.delenv("PPTX_MCP_API_KEY", raising=False)
    client = TestClient(_create_app())

    response_health = client.get("/health")
    response_healthz = client.get("/healthz")

    assert response_health.status_code == 200
    assert response_healthz.status_code == 200


def test_options_passthrough(monkeypatch):
    monkeypatch.delenv("DOC_DOWNLOAD_SIGNING_SECRET", raising=False)
    monkeypatch.delenv("PPTX_MCP_API_KEY", raising=False)
    client = TestClient(_create_app())

    response = client.options("/protected")

    assert response.status_code == 200


def test_signed_file_url_valid_without_api_key(monkeypatch):
    monkeypatch.setenv("DOC_DOWNLOAD_SIGNING_SECRET", "signing-secret")
    monkeypatch.delenv("PPTX_MCP_API_KEY", raising=False)
    client = TestClient(_create_app())

    expires_at = int(time.time()) + 120
    signature = build_download_signature("sample.pptx", expires_at, "signing-secret")
    response = client.get(f"/files/sample.pptx?exp={expires_at}&sig={signature}")

    assert response.status_code == 200
    assert response.json() == {"ok": True}


def test_signed_file_url_alias_params_are_accepted(monkeypatch):
    monkeypatch.setenv("DOC_DOWNLOAD_SIGNING_SECRET", "signing-secret")
    monkeypatch.delenv("PPTX_MCP_API_KEY", raising=False)
    client = TestClient(_create_app())

    expires_at = int(time.time()) + 120
    signature = build_download_signature("sample.pptx", expires_at, "signing-secret")
    response = client.get(f"/files/sample.pptx?expires={expires_at}&signature={signature}")

    assert response.status_code == 200


def test_invalid_or_expired_signature_returns_403_even_with_api_key(monkeypatch):
    monkeypatch.setenv("DOC_DOWNLOAD_SIGNING_SECRET", "signing-secret")
    monkeypatch.delenv("PPTX_MCP_API_KEY", raising=False)
    client = TestClient(_create_app())

    expires_at = int(time.time()) + 120
    response = client.get(
        f"/files/sample.pptx?exp={expires_at}&sig=bad-signature",
        headers={"x-api-key": "expected-secret"},
    )

    assert response.status_code == 403
    assert response.json() == {"error": "Invalid or expired download signature."}


def test_expired_signature_returns_json_by_default(monkeypatch):
    monkeypatch.setenv("DOC_DOWNLOAD_SIGNING_SECRET", "signing-secret")
    monkeypatch.delenv("PPTX_MCP_API_KEY", raising=False)
    client = TestClient(_create_app())

    expires_at = int(time.time()) - 5
    signature = build_download_signature("sample.pptx", expires_at, "signing-secret")
    response = client.get(f"/files/sample.pptx?exp={expires_at}&sig={signature}")

    assert response.status_code == 403
    assert response.json() == {"error": "Invalid or expired download signature."}


def test_expired_signature_returns_html_for_browser_requests(monkeypatch):
    monkeypatch.setenv("DOC_DOWNLOAD_SIGNING_SECRET", "signing-secret")
    monkeypatch.delenv("PPTX_MCP_API_KEY", raising=False)
    client = TestClient(_create_app())

    expires_at = int(time.time()) - 5
    signature = build_download_signature("sample.pptx", expires_at, "signing-secret")
    response = client.get(
        f"/files/sample.pptx?exp={expires_at}&sig={signature}",
        headers={"accept": "text/html"},
    )

    assert response.status_code == 403
    assert "text/html" in response.headers.get("content-type", "")
    assert "Download Link Expired" in response.text


def test_file_route_without_signature_still_requires_api_key(monkeypatch):
    monkeypatch.setenv("DOC_DOWNLOAD_SIGNING_SECRET", "signing-secret")
    monkeypatch.delenv("PPTX_MCP_API_KEY", raising=False)
    client = TestClient(_create_app())

    response = client.get("/files/sample.pptx")

    assert response.status_code == 401
    assert response.json() == {"error": "Missing API key header: x-api-key"}


def test_build_download_url_signing_and_ttl(monkeypatch):
    monkeypatch.setenv("DOC_DOWNLOAD_SIGNING_SECRET", "signing-secret")
    monkeypatch.delenv("PPTX_MCP_API_KEY", raising=False)
    monkeypatch.setattr(http_auth.time, "time", lambda: 1000)
    monkeypatch.delenv("DOC_DOWNLOAD_URL_TTL_SECONDS", raising=False)

    url = http_auth.build_download_url("https://example.com/files", "sample.pptx")
    parsed = urlparse(url)
    params = parse_qs(parsed.query)

    assert parsed.path == "/files/sample.pptx"
    assert "exp" in params
    assert params["exp"][0] == "1900"
    assert "sig" in params


def test_build_download_url_uses_minimum_ttl(monkeypatch):
    monkeypatch.setenv("DOC_DOWNLOAD_SIGNING_SECRET", "signing-secret")
    monkeypatch.delenv("PPTX_MCP_API_KEY", raising=False)
    monkeypatch.setenv("DOC_DOWNLOAD_URL_TTL_SECONDS", "5")
    monkeypatch.setattr(http_auth.time, "time", lambda: 1000)

    url = http_auth.build_download_url("https://example.com/files", "sample.pptx")
    params = parse_qs(urlparse(url).query)

    assert params["exp"][0] == "1030"


def test_build_download_url_falls_back_to_api_key_secret(monkeypatch):
    monkeypatch.delenv("DOC_DOWNLOAD_SIGNING_SECRET", raising=False)
    monkeypatch.setenv("PPTX_MCP_API_KEY", "api-key-secret")
    monkeypatch.setattr(http_auth.time, "time", lambda: 1000)

    url = http_auth.build_download_url("https://example.com/files", "sample.pptx")
    params = parse_qs(urlparse(url).query)
    expected_sig = build_download_signature("sample.pptx", 1900, "api-key-secret")

    assert params["sig"][0] == expected_sig


def test_build_download_url_returns_plain_url_without_secret(monkeypatch):
    monkeypatch.delenv("DOC_DOWNLOAD_SIGNING_SECRET", raising=False)
    monkeypatch.delenv("PPTX_MCP_API_KEY", raising=False)

    url = http_auth.build_download_url("https://example.com/files", "sample.pptx")

    assert url == "https://example.com/files/sample.pptx"


class _DummyTransportApp:
    def __init__(self):
        self.settings = SimpleNamespace(
            host=None,
            port=None,
            sse_path="/events",
            transport_security=SimpleNamespace(allowed_hosts=[]),
        )
        self.run_calls = []

    def run(self, transport: str):
        self.run_calls.append(transport)


class _DummyASGIApp:
    def __init__(self):
        self.middleware_calls = []

    def add_middleware(self, middleware_cls, **kwargs):
        self.middleware_calls.append((middleware_cls, kwargs))


def test_build_sse_app_prefers_http_app(monkeypatch):
    sentinel = object()

    class _FakeApp:
        settings = SimpleNamespace(sse_path="/events")

        def http_app(self, transport=None):
            assert transport == "sse"
            return sentinel

    monkeypatch.setattr(ppt_mcp_server, "app", _FakeApp())

    assert ppt_mcp_server._build_sse_app() is sentinel


def test_build_sse_app_falls_back_to_sse_app_with_path(monkeypatch):
    sentinel = object()
    calls = []

    class _FakeApp:
        settings = SimpleNamespace(sse_path="/events")

        def http_app(self, transport=None):
            raise TypeError("transport arg unsupported")

        def sse_app(self, path=None):
            calls.append(path)
            return sentinel

    monkeypatch.setattr(ppt_mcp_server, "app", _FakeApp())

    assert ppt_mcp_server._build_sse_app() is sentinel
    assert calls == ["/events"]


def test_build_sse_app_falls_back_to_sse_app_without_path(monkeypatch):
    sentinel = object()
    calls = []

    class _FakeApp:
        settings = SimpleNamespace(sse_path="/events")

        def http_app(self, transport=None):
            raise TypeError("transport arg unsupported")

        def sse_app(self, path=None):
            calls.append(path)
            if path is not None:
                raise TypeError("path arg unsupported")
            return sentinel

    monkeypatch.setattr(ppt_mcp_server, "app", _FakeApp())

    assert ppt_mcp_server._build_sse_app() is sentinel
    assert calls == ["/events", None]


def test_build_sse_app_raises_when_unavailable(monkeypatch):
    class _FakeApp:
        settings = SimpleNamespace(sse_path="/events")

    monkeypatch.setattr(ppt_mcp_server, "app", _FakeApp())

    with pytest.raises(RuntimeError, match="SSE app builder"):
        ppt_mcp_server._build_sse_app()


def test_main_sse_with_api_key_uses_middleware_backed_asgi(monkeypatch):
    dummy_transport_app = _DummyTransportApp()
    dummy_asgi_app = _DummyASGIApp()
    uvicorn_calls = {}

    monkeypatch.setattr(ppt_mcp_server, "app", dummy_transport_app)
    monkeypatch.setattr(ppt_mcp_server, "_build_sse_app", lambda: dummy_asgi_app)
    monkeypatch.setenv("PPTX_MCP_API_KEY", "expected-secret")
    monkeypatch.setenv("PPTX_MCP_API_KEY_HEADER", "x-custom-key")
    monkeypatch.setenv("PORT", "8123")
    monkeypatch.delenv("RAILWAY_PUBLIC_DOMAIN", raising=False)
    monkeypatch.delenv("RAILWAY_PRIVATE_DOMAIN", raising=False)

    fake_uvicorn = SimpleNamespace(
        run=lambda app, host, port: uvicorn_calls.update(
            {"app": app, "host": host, "port": port}
        )
    )
    monkeypatch.setitem(sys.modules, "uvicorn", fake_uvicorn)

    ppt_mcp_server.main("sse")

    assert uvicorn_calls["app"] is dummy_asgi_app
    assert uvicorn_calls["host"] == "0.0.0.0"
    assert uvicorn_calls["port"] == 8123
    assert dummy_transport_app.run_calls == []
    assert len(dummy_asgi_app.middleware_calls) == 1
    middleware_cls, kwargs = dummy_asgi_app.middleware_calls[0]
    assert middleware_cls is APIKeyMiddleware
    assert kwargs["api_key"] == "expected-secret"
    assert kwargs["header_name"] == "x-custom-key"
    assert kwargs["exempt_paths"] == ["/health", "/healthz"]


def test_main_sse_with_api_key_missing_builder_fails_fast(monkeypatch):
    dummy_transport_app = _DummyTransportApp()

    monkeypatch.setattr(ppt_mcp_server, "app", dummy_transport_app)
    monkeypatch.setattr(
        ppt_mcp_server,
        "_build_sse_app",
        lambda: (_ for _ in ()).throw(RuntimeError("no builder")),
    )
    monkeypatch.setenv("PPTX_MCP_API_KEY", "expected-secret")
    monkeypatch.setenv("PORT", "8123")

    fake_uvicorn = SimpleNamespace(run=lambda app, host, port: None)
    monkeypatch.setitem(sys.modules, "uvicorn", fake_uvicorn)

    with pytest.raises(RuntimeError, match="requires ASGI SSE app builder support"):
        ppt_mcp_server.main("sse")

    assert dummy_transport_app.run_calls == []


def test_main_sse_without_api_key_uses_legacy_run(monkeypatch):
    dummy_transport_app = _DummyTransportApp()

    monkeypatch.setattr(ppt_mcp_server, "app", dummy_transport_app)
    monkeypatch.delenv("PPTX_MCP_API_KEY", raising=False)
    monkeypatch.delenv("PORT", raising=False)
    monkeypatch.delenv("RAILWAY_PUBLIC_DOMAIN", raising=False)
    monkeypatch.delenv("RAILWAY_PRIVATE_DOMAIN", raising=False)

    ppt_mcp_server.main("sse")

    assert dummy_transport_app.run_calls == ["sse"]

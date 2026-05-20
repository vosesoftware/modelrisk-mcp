"""Tests for the bearer-token middleware.

We boot a tiny Starlette app, wrap it with the middleware, and hit it
with a synchronous test client to confirm:
- missing / malformed header → 401
- wrong token → 401
- correct token → reaches the handler
- empty configured token → rejected at build time
"""

from __future__ import annotations

import pytest
from starlette.applications import Starlette
from starlette.responses import PlainTextResponse
from starlette.routing import Route
from starlette.testclient import TestClient

from modelrisk_mcp.http_auth import build_auth_middleware


def _handler(_request):  # type: ignore[no-untyped-def]
    return PlainTextResponse("ok")


def _build_app(token: str) -> Starlette:
    app = Starlette(routes=[Route("/mcp", _handler, methods=["GET", "POST"])])
    app.add_middleware(build_auth_middleware(token))
    return app


class TestBearerTokenMiddleware:
    def test_missing_header_rejected(self) -> None:
        client = TestClient(_build_app("secret"))
        response = client.post("/mcp")
        assert response.status_code == 401
        body = response.json()
        assert body["error"] == "unauthorised"
        assert "missing" in body["reason"]
        assert response.headers["WWW-Authenticate"].startswith("Bearer")

    def test_malformed_header_rejected(self) -> None:
        client = TestClient(_build_app("secret"))
        response = client.post("/mcp", headers={"Authorization": "Basic abc"})
        assert response.status_code == 401

    def test_wrong_token_rejected(self) -> None:
        client = TestClient(_build_app("secret"))
        response = client.post("/mcp", headers={"Authorization": "Bearer wrong"})
        assert response.status_code == 401
        assert response.json()["reason"] == "invalid bearer token"

    def test_correct_token_passes_through(self) -> None:
        client = TestClient(_build_app("secret"))
        response = client.post(
            "/mcp", headers={"Authorization": "Bearer secret"}
        )
        assert response.status_code == 200
        assert response.text == "ok"

    def test_bearer_keyword_is_case_insensitive(self) -> None:
        client = TestClient(_build_app("secret"))
        response = client.post(
            "/mcp", headers={"Authorization": "bearer secret"}
        )
        # Bearer auth headers are case-insensitive per RFC 6750.
        assert response.status_code == 200

    def test_empty_token_build_rejected(self) -> None:
        with pytest.raises(ValueError):
            build_auth_middleware("")

    def test_token_with_internal_whitespace(self) -> None:
        """A bearer token containing whitespace is unusual but should
        compare exactly. We trim only outer whitespace from the header
        value, not from inside the token."""
        client = TestClient(_build_app("a b c"))
        response = client.post(
            "/mcp", headers={"Authorization": "Bearer a b c"}
        )
        assert response.status_code == 200

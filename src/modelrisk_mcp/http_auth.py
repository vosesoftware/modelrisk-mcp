"""Bearer-token middleware for the HTTP transports.

Used when `--token` (or `MODELRISK_MCP_TOKEN`) is set. Requests without a
valid `Authorization: Bearer <token>` header are rejected with a 401
that doesn't reveal the expected token's length.

The middleware is built around `BaseHTTPMiddleware` from Starlette
because FastMCP returns a Starlette app for both `streamable-http` and
`sse`. Constant-time comparison via `secrets.compare_digest` blocks
timing-based token extraction.
"""

from __future__ import annotations

import hmac
from collections.abc import Awaitable, Callable

from starlette.middleware.base import BaseHTTPMiddleware
from starlette.requests import Request
from starlette.responses import JSONResponse, Response


class _BearerTokenMiddleware(BaseHTTPMiddleware):
    """Starlette middleware enforcing `Authorization: Bearer <token>`."""

    _expected_token: str

    async def dispatch(
        self,
        request: Request,
        call_next: Callable[[Request], Awaitable[Response]],
    ) -> Response:
        header = request.headers.get("authorization", "")
        if not header.lower().startswith("bearer "):
            return _unauthorised("missing or malformed Authorization header")
        supplied = header.split(None, 1)[1].strip()
        if not hmac.compare_digest(supplied, self._expected_token):
            return _unauthorised("invalid bearer token")
        return await call_next(request)


def _unauthorised(reason: str) -> JSONResponse:
    return JSONResponse(
        {"error": "unauthorised", "reason": reason},
        status_code=401,
        headers={"WWW-Authenticate": 'Bearer realm="modelrisk-mcp"'},
    )


def build_auth_middleware(token: str) -> type[_BearerTokenMiddleware]:
    """Build a middleware class bound to the given expected token.

    Returning a class (rather than an instance) lets Starlette's
    `app.add_middleware(MiddlewareClass)` API work cleanly."""
    if not token:
        raise ValueError("token must be non-empty")

    class _BoundMiddleware(_BearerTokenMiddleware):
        _expected_token = token

    _BoundMiddleware.__name__ = "BearerTokenMiddleware"
    return _BoundMiddleware

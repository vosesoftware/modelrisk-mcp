"""ModelRisk MCP server entrypoint.

Supports three MCP transports per the spec — stdio (default) for
Claude Desktop / Code / Cursor / Zed, and streamable-http or sse for
Claude for Excel and other remote MCP clients that can't spawn local
subprocesses.

CLI:

    python -m modelrisk_mcp                                  # stdio (default)
    python -m modelrisk_mcp --transport=streamable-http      # HTTP, 127.0.0.1:8000
    python -m modelrisk_mcp --transport=sse --port=9000
    python -m modelrisk_mcp --transport=streamable-http --token=$(uuid)

A bearer token is strongly recommended for HTTP transports: any process
on the same machine can otherwise hit the endpoint and drive Excel. The
token is set via `--token`, the `MODELRISK_MCP_TOKEN` environment
variable, or omitted (open access — only safe for `127.0.0.1` development).
"""

from __future__ import annotations

import argparse
import os
import sys
from typing import Literal

from modelrisk_mcp.server import mcp

Transport = Literal["stdio", "streamable-http", "sse"]
_VALID_TRANSPORTS: tuple[Transport, ...] = ("stdio", "streamable-http", "sse")


def _build_parser() -> argparse.ArgumentParser:
    p = argparse.ArgumentParser(
        prog="modelrisk-mcp",
        description=(
            "ModelRisk MCP server — exposes ModelRisk's read/build/run "
            "surface to Claude Desktop, Claude Code, Claude for Excel, "
            "Cursor, Zed, and any MCP-compliant client."
        ),
    )
    p.add_argument(
        "--transport",
        choices=_VALID_TRANSPORTS,
        default="stdio",
        help=(
            "Transport protocol. 'stdio' (default) is for local clients "
            "like Claude Desktop; 'streamable-http' is the modern remote "
            "MCP transport (recommended for Claude for Excel); 'sse' is "
            "the legacy SSE transport."
        ),
    )
    p.add_argument(
        "--host",
        default="127.0.0.1",
        help=(
            "Bind host for HTTP transports. Defaults to 127.0.0.1 "
            "(loopback only) — change to 0.0.0.0 only if you understand "
            "the security implications."
        ),
    )
    p.add_argument(
        "--port",
        type=int,
        default=8000,
        help="Bind port for HTTP transports. Default: 8000.",
    )
    p.add_argument(
        "--mount-path",
        default=None,
        help=(
            "Mount path for the MCP endpoint on HTTP transports. "
            "Default: '/mcp' for streamable-http, '/sse' for SSE."
        ),
    )
    p.add_argument(
        "--token",
        default=None,
        help=(
            "Bearer token required on HTTP transports (sent as "
            "'Authorization: Bearer <token>'). Falls back to the "
            "MODELRISK_MCP_TOKEN environment variable if unset. Strongly "
            "recommended for any non-loopback HTTP deployment."
        ),
    )
    return p


def _run_stdio() -> None:
    mcp.run(transport="stdio")


def _run_http(
    *,
    transport: Literal["streamable-http", "sse"],
    host: str,
    port: int,
    mount_path: str | None,
    token: str | None,
) -> None:
    try:
        import uvicorn
    except ImportError as exc:  # pragma: no cover — dependency is declared
        raise SystemExit(
            "uvicorn is required for HTTP transports; install with "
            "`uv sync` or `pip install modelrisk-mcp[dev]`."
        ) from exc

    from modelrisk_mcp.http_auth import build_auth_middleware

    if transport == "streamable-http":
        app = mcp.streamable_http_app()
    else:
        app = mcp.sse_app(mount_path=mount_path) if mount_path else mcp.sse_app()

    effective_token = token or os.environ.get("MODELRISK_MCP_TOKEN")
    if effective_token:
        app.add_middleware(build_auth_middleware(effective_token))
    elif host != "127.0.0.1":
        print(
            "WARNING: starting HTTP transport on a non-loopback host without "
            "a bearer token. Any process on the network can drive Excel. "
            "Set --token or MODELRISK_MCP_TOKEN.",
            file=sys.stderr,
        )

    uvicorn.run(app, host=host, port=port, log_level="info")


def main(argv: list[str] | None = None) -> None:
    args = _build_parser().parse_args(argv)
    if args.transport == "stdio":
        _run_stdio()
    else:
        _run_http(
            transport=args.transport,
            host=args.host,
            port=args.port,
            mount_path=args.mount_path,
            token=args.token,
        )


if __name__ == "__main__":
    main()

"""ModelRisk MCP server entrypoint.

Supports three MCP transports per the spec — stdio (default) for
Claude Desktop / Code / Cursor / Zed, and streamable-http or sse for
MCP clients on your own machine that speak HTTP rather than spawning a
subprocess. (Claude for Excel cannot reach a local server at all — see
docs/claude-for-excel.md.)

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
from typing import Literal, cast

from modelrisk_mcp.server import mcp

Transport = Literal["stdio", "streamable-http", "sse"]
_VALID_TRANSPORTS: tuple[Transport, ...] = ("stdio", "streamable-http", "sse")


def _build_parser() -> argparse.ArgumentParser:
    p = argparse.ArgumentParser(
        prog="modelrisk-mcp",
        description=(
            "ModelRisk MCP server — exposes ModelRisk's read/build/run "
            "surface to Claude Desktop, Claude Code, Cursor, Zed, and any "
            "local MCP-compliant client."
        ),
    )
    p.add_argument(
        "--transport",
        choices=_VALID_TRANSPORTS,
        default="stdio",
        help=(
            "Transport protocol. 'stdio' (default) is for local clients "
            "like Claude Desktop; 'streamable-http' is the modern remote "
            "MCP transport for clients running on your own machine; 'sse' is "
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
    p.add_argument(
        "--read-only",
        action="store_true",
        help=(
            "Disable every mutating operation: cell/range writes, named "
            "ranges, simulations, and workbook saves all raise a clear "
            "error while the reading and analysis tools keep working. "
            "The recommended first-session posture when pointing the "
            "server at a live client model. Equivalent to setting "
            "MODELRISK_MCP_READ_ONLY=1."
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

    # mcp 2.0 renamed the path kwargs (1.x `mount_path` is gone) and
    # added Host-header validation configured via `host` — pass the
    # bind host so non-loopback deployments validate correctly.
    if transport == "streamable-http":
        app = mcp.streamable_http_app(
            streamable_http_path=mount_path or "/mcp", host=host,
        )
    else:
        app = mcp.sse_app(sse_path=mount_path or "/sse", host=host)

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
    if argv is None:
        argv = sys.argv[1:]

    # Subcommand dispatch — we keep the legacy "no subcommand = serve"
    # behaviour because Claude Desktop configs in the wild say
    # `"command": "modelrisk-mcp"` with no args. Only intercept the
    # first arg if it's a recognised non-server subcommand.
    if argv and argv[0] in {"install", "uninstall"}:
        return _run_install(
            cast(Literal["install", "uninstall"], argv[0]), argv[1:]
        )
    if argv and argv[0] == "serve":
        argv = argv[1:]  # `modelrisk-mcp serve --transport=stdio` works too

    args = _build_parser().parse_args(argv)
    if args.read_only:
        # The enforcement layer (ExcelBridge._ensure_writable) checks
        # this env var at call time, so setting it here covers every
        # transport and entry path.
        import os

        os.environ["MODELRISK_MCP_READ_ONLY"] = "1"
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


def _run_install(mode: Literal["install", "uninstall"], argv: list[str]) -> None:
    """Dispatch the install / uninstall subcommand.

    Kept in `__main__.py` (rather than `install.py`) because it's the
    CLI argparse glue — `install.py` is import-safe and side-effect-
    free."""
    from modelrisk_mcp import install as install_mod

    parser = argparse.ArgumentParser(
        prog=f"modelrisk-mcp {mode}",
        description=(
            "Add the modelrisk MCP server entry to every detected MCP "
            "client's config (Claude Desktop, Claude Code). Backs up "
            "existing configs before writing."
            if mode == "install"
            else
            "Remove the modelrisk MCP server entry from every detected "
            "MCP client's config. Idempotent."
        ),
    )
    parser.add_argument(
        "--name",
        default="modelrisk",
        help=(
            "Server name under `mcpServers`. Default: 'modelrisk'. Useful "
            "if you want to register multiple instances (e.g. dev + "
            "production) side by side."
        ),
    )
    if mode == "install":
        parser.add_argument(
            "--force",
            action="store_true",
            help=(
                "Overwrite an existing entry with the same name. Without "
                "this flag, install skips clients where the name is "
                "already taken by a different command."
            ),
        )
    args = parser.parse_args(argv)

    try:
        if mode == "install":
            results = install_mod.install(
                server_name=args.name, force=args.force,
            )
        else:
            results = install_mod.uninstall(server_name=args.name)
    except install_mod.InstallError as exc:
        print(f"error: {exc}", file=sys.stderr)
        raise SystemExit(1) from exc

    exit_code = 0
    for r in results:
        symbol = {
            "added":   "+",
            "updated": "~",
            "removed": "-",
            "skipped": ".",
            "error":   "X",
        }.get(r.action, "?")
        print(f"  {symbol} {r.client:<16} {r.action:<8} {r.config_path}")
        if r.message:
            print(f"      {r.message}")
        if r.backup_path:
            print(f"      backup: {r.backup_path}")
        if r.action == "error":
            exit_code = 1
    if mode == "install":
        print()
        print("Restart Claude Desktop / Claude Code to pick up the new server.")
    raise SystemExit(exit_code)


if __name__ == "__main__":
    main()

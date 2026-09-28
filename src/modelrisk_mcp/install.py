"""Wire `modelrisk-mcp` into MCP-client config files automatically.

End users typically install a Python package then hand-edit Claude
Desktop's `claude_desktop_config.json` to register the server. The
JSON-editing step is the worst part of MCP setup — wrong key names,
missing commas, accidental clobbering of other servers' entries.

`modelrisk-mcp install` does the wiring on the user's behalf:

- Detects which MCP clients are installed (Claude Desktop, Claude Code).
- Backs up each existing config.
- Merges in the `modelrisk` server entry (preserves any other servers).
- Picks the right `command` path: the absolute path to the installed
  `modelrisk-mcp.exe` (so the user doesn't need to worry whether
  `Scripts/` is on PATH for Claude's subprocess).

The reverse: `modelrisk-mcp install --uninstall` removes our entry
without touching anything else.
"""

from __future__ import annotations

import json
import os
import shutil
import sys
import time
from dataclasses import dataclass
from pathlib import Path
from typing import Any

# Default server entry key. Users can override with --name; this is the
# name that appears as Claude's "ModelRisk tools" section header.
_DEFAULT_SERVER_NAME = "modelrisk"


# ---------------------------------------------------------------------------
# Client discovery
# ---------------------------------------------------------------------------


@dataclass(frozen=True)
class ClientTarget:
    """One MCP client we know how to register a server with."""

    name: str
    config_path: Path
    # JSON key holding the server map. Claude Desktop uses top-level
    # `mcpServers` in claude_desktop_config.json; Claude Code uses
    # top-level `mcpServers` in ~/.claude.json.
    servers_key: str = "mcpServers"
    # "json" = merge into config_path directly. "claude-cli" = prefer
    # shelling out to `claude mcp add/remove` (the CLI owns the schema),
    # falling back to a direct JSON merge into config_path.
    strategy: str = "json"


def _claude_desktop_config_path() -> Path:
    """Windows: %APPDATA%\\Claude\\claude_desktop_config.json.
    macOS: ~/Library/Application Support/Claude/claude_desktop_config.json.
    Linux: not officially supported by Claude Desktop."""
    if sys.platform == "win32":
        appdata = os.environ.get("APPDATA")
        if appdata:
            return Path(appdata) / "Claude" / "claude_desktop_config.json"
    elif sys.platform == "darwin":
        return (
            Path.home()
            / "Library"
            / "Application Support"
            / "Claude"
            / "claude_desktop_config.json"
        )
    # Fallback for development on other platforms — best effort.
    return Path.home() / ".config" / "Claude" / "claude_desktop_config.json"


def _claude_code_config_path() -> Path:
    """Claude Code reads user-scope MCP servers from the TOP-LEVEL
    `mcpServers` key of `~/.claude.json` — NOT `~/.claude/settings.json`,
    which holds per-user settings (theme, notifications, …) and is
    silently ignored for MCP registration. Writing there produced a
    green 'added' line and a dead server (field bug report,
    2026-07-20)."""
    return Path.home() / ".claude.json"


def _claude_cli() -> str | None:
    """Path to the Claude Code CLI, if installed."""
    return shutil.which("claude")


def discover_clients() -> list[ClientTarget]:
    """Return only the clients that actually appear to be installed.

    Claude Desktop: its config directory exists.
    Claude Code: `~/.claude.json` exists OR the `claude` CLI is on PATH
    (the old check — `~/.claude/` existing — mis-detected Claude Code
    via a directory other tools also create)."""
    targets: list[ClientTarget] = []
    desktop = ClientTarget("Claude Desktop", _claude_desktop_config_path())
    if desktop.config_path.parent.is_dir():
        targets.append(desktop)
    code_cfg = _claude_code_config_path()
    if code_cfg.is_file() or _claude_cli():
        targets.append(
            ClientTarget("Claude Code", code_cfg, strategy="claude-cli")
        )
    return targets


# ---------------------------------------------------------------------------
# Server entry composition
# ---------------------------------------------------------------------------


def resolve_server_entry() -> dict[str, Any]:
    """Return the JSON object that should be written for our server.

    Strategy:
    1. Running as the standalone exe (PyInstaller sets `sys.frozen`):
       register that exe, with no args. There `sys.executable` is the
       exe itself, which takes no interpreter flags, so fallback 3 made
       the registered server exit 2 at every start. It also wins over a
       pip-installed `modelrisk-mcp` on PATH: install registers the
       program you ran.
    2. Prefer the absolute path to the installed `modelrisk-mcp` exe
       found on PATH. Most robust against PATH not being set up for
       Claude's spawned subprocess.
    3. Fall back to running this Python interpreter as `-m modelrisk_mcp`.
       Works regardless of installed-shim status.
    """
    if getattr(sys, "frozen", False):
        return {"command": sys.executable}
    exe = shutil.which("modelrisk-mcp")
    if exe:
        return {"command": exe}
    return {"command": sys.executable, "args": ["-m", "modelrisk_mcp"]}


# ---------------------------------------------------------------------------
# Config IO
# ---------------------------------------------------------------------------


def _read_config(path: Path) -> dict[str, Any]:
    if not path.is_file():
        return {}
    try:
        text = path.read_text(encoding="utf-8")
    except OSError as exc:
        raise InstallError(f"Cannot read {path}: {exc}") from exc
    text = text.strip()
    if not text:
        return {}
    try:
        data = json.loads(text)
    except json.JSONDecodeError as exc:
        raise InstallError(
            f"{path} is not valid JSON: {exc}. Fix the file by hand "
            "first, then re-run install."
        ) from exc
    if not isinstance(data, dict):
        raise InstallError(f"{path} top-level value must be a JSON object.")
    return data


def _backup(path: Path) -> Path | None:
    if not path.is_file():
        return None
    stamp = time.strftime("%Y%m%d-%H%M%S")
    backup = path.with_suffix(path.suffix + f".bak.{stamp}")
    shutil.copy2(path, backup)
    return backup


def _write_config_atomic(path: Path, data: dict[str, Any]) -> None:
    """Write JSON via temp-file + rename so a crashed/interrupted
    write can't leave a half-written config."""
    path.parent.mkdir(parents=True, exist_ok=True)
    tmp = path.with_suffix(path.suffix + ".tmp")
    tmp.write_text(json.dumps(data, indent=2) + "\n", encoding="utf-8")
    os.replace(tmp, path)


# ---------------------------------------------------------------------------
# Public API
# ---------------------------------------------------------------------------


class InstallError(RuntimeError):
    """Surface clean error messages without raw stack traces."""


@dataclass
class InstallResult:
    """What happened to one client target."""

    client: str
    config_path: Path
    action: str   # "added", "updated", "removed", "skipped", "error"
    backup_path: Path | None = None
    message: str = ""


def install(
    *,
    server_name: str = _DEFAULT_SERVER_NAME,
    clients: list[ClientTarget] | None = None,
    server_entry: dict[str, Any] | None = None,
    force: bool = False,
) -> list[InstallResult]:
    """Add the modelrisk server entry to every discovered client config.

    Returns a per-client result list. Does not raise on per-client
    errors (returns them in the result) so a partial success still
    reports cleanly. Raises `InstallError` only on completely fatal
    conditions (e.g. no clients found).
    """
    targets = clients if clients is not None else discover_clients()
    if not targets:
        raise InstallError(
            "No supported MCP clients detected on this machine. Supported: "
            "Claude Desktop (%APPDATA%/Claude/), Claude Code (~/.claude/)."
        )
    entry = server_entry if server_entry is not None else resolve_server_entry()
    results: list[InstallResult] = []
    for target in targets:
        try:
            results.append(
                _install_one(target, server_name, entry, force=force)
            )
        except InstallError as exc:
            results.append(
                InstallResult(
                    client=target.name,
                    config_path=target.config_path,
                    action="error",
                    message=str(exc),
                )
            )
    return results


def uninstall(
    *,
    server_name: str = _DEFAULT_SERVER_NAME,
    clients: list[ClientTarget] | None = None,
) -> list[InstallResult]:
    """Remove the modelrisk server entry from every discovered client.
    Idempotent: if the entry isn't there, we report 'skipped' rather
    than error."""
    targets = clients if clients is not None else discover_clients()
    if not targets:
        raise InstallError("No supported MCP clients detected.")
    results: list[InstallResult] = []
    for target in targets:
        try:
            results.append(_uninstall_one(target, server_name))
        except InstallError as exc:
            results.append(
                InstallResult(
                    client=target.name,
                    config_path=target.config_path,
                    action="error",
                    message=str(exc),
                )
            )
    return results


def _claude_cli_add(
    server_name: str, entry: dict[str, Any], *, force: bool
) -> InstallResult | None:
    """Register via `claude mcp add` — the CLI owns the config schema
    and scoping, so prefer it whenever it's on PATH. Returns None when
    the CLI is unavailable (caller falls back to the JSON merge)."""
    import subprocess

    cli = _claude_cli()
    if not cli:
        return None
    command = entry["command"]
    args = entry.get("args", [])
    if force:
        # Idempotent replace: remove any existing entry first; ignore
        # 'not found' failures.
        subprocess.run(
            [cli, "mcp", "remove", server_name, "--scope", "user"],
            capture_output=True, text=True, timeout=30,
        )
    proc = subprocess.run(
        [cli, "mcp", "add", server_name, "--scope", "user", "--",
         command, *args],
        capture_output=True, text=True, timeout=30,
    )
    if proc.returncode != 0:
        detail = (proc.stderr or proc.stdout or "").strip()
        if "already exists" in detail.lower() and not force:
            return InstallResult(
                client="Claude Code",
                config_path=_claude_code_config_path(),
                action="skipped",
                message=(
                    f"{server_name!r} already registered in Claude Code. "
                    "Re-run with --force to replace it."
                ),
            )
        raise InstallError(
            f"`claude mcp add` failed ({proc.returncode}): {detail[:300]}"
        )
    return InstallResult(
        client="Claude Code",
        config_path=_claude_code_config_path(),
        action="added",
        message=(
            f"Registered {server_name!r} via `claude mcp add --scope user` "
            f"-> {command} {' '.join(args)}".rstrip()
        ),
    )


def _claude_cli_remove(server_name: str) -> InstallResult | None:
    import subprocess

    cli = _claude_cli()
    if not cli:
        return None
    proc = subprocess.run(
        [cli, "mcp", "remove", server_name, "--scope", "user"],
        capture_output=True, text=True, timeout=30,
    )
    if proc.returncode != 0:
        detail = (proc.stderr or proc.stdout or "").strip().lower()
        if "not found" in detail or "no mcp server" in detail:
            return InstallResult(
                client="Claude Code",
                config_path=_claude_code_config_path(),
                action="skipped",
                message=f"{server_name!r} not registered in Claude Code.",
            )
        raise InstallError(
            f"`claude mcp remove` failed ({proc.returncode}): {detail[:300]}"
        )
    return InstallResult(
        client="Claude Code",
        config_path=_claude_code_config_path(),
        action="removed",
        message=f"Removed {server_name!r} via `claude mcp remove`.",
    )


def _install_one(
    target: ClientTarget,
    server_name: str,
    entry: dict[str, Any],
    *,
    force: bool,
) -> InstallResult:
    if target.strategy == "claude-cli":
        via_cli = _claude_cli_add(server_name, entry, force=force)
        if via_cli is not None:
            return via_cli
        # CLI absent — fall through to a direct merge into ~/.claude.json.
    data = _read_config(target.config_path)
    servers = data.setdefault(target.servers_key, {})
    if not isinstance(servers, dict):
        raise InstallError(
            f"{target.config_path}: '{target.servers_key}' is not an object."
        )

    existing = servers.get(server_name)
    if existing is not None and existing != entry and not force:
        return InstallResult(
            client=target.name,
            config_path=target.config_path,
            action="skipped",
            message=(
                f"{server_name!r} already configured with a different command. "
                "Re-run with --force to overwrite, or use --name <other> to "
                "register under a different key."
            ),
        )

    backup = _backup(target.config_path)
    action = "updated" if existing is not None else "added"
    servers[server_name] = entry
    _write_config_atomic(target.config_path, data)
    return InstallResult(
        client=target.name,
        config_path=target.config_path,
        action=action,
        backup_path=backup,
        message=f"Registered {server_name!r} -> {entry}",
    )


def _uninstall_one(target: ClientTarget, server_name: str) -> InstallResult:
    if target.strategy == "claude-cli":
        via_cli = _claude_cli_remove(server_name)
        if via_cli is not None:
            return via_cli
    if not target.config_path.is_file():
        return InstallResult(
            client=target.name,
            config_path=target.config_path,
            action="skipped",
            message="config file does not exist.",
        )
    data = _read_config(target.config_path)
    servers = data.get(target.servers_key, {})
    if not isinstance(servers, dict) or server_name not in servers:
        return InstallResult(
            client=target.name,
            config_path=target.config_path,
            action="skipped",
            message=f"{server_name!r} not present in config.",
        )
    backup = _backup(target.config_path)
    del servers[server_name]
    _write_config_atomic(target.config_path, data)
    return InstallResult(
        client=target.name,
        config_path=target.config_path,
        action="removed",
        backup_path=backup,
        message=f"Removed {server_name!r}.",
    )


__all__ = [
    "ClientTarget",
    "InstallError",
    "InstallResult",
    "discover_clients",
    "install",
    "resolve_server_entry",
    "uninstall",
]

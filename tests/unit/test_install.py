"""Tests for the `modelrisk-mcp install` machinery.

We test against tmp_path-rooted fake config files rather than the
real Claude Desktop config — same logic, no risk of clobbering the
developer's actual setup.
"""

from __future__ import annotations

import json
from pathlib import Path

import pytest

from modelrisk_mcp.install import (
    ClientTarget,
    InstallError,
    install,
    resolve_server_entry,
    uninstall,
)


def _make_target(tmp_path: Path, name: str = "Claude Desktop") -> ClientTarget:
    cfg = tmp_path / "fake-client" / "config.json"
    cfg.parent.mkdir(parents=True, exist_ok=True)
    return ClientTarget(name=name, config_path=cfg)


# ----------------------------------------------------------------------
# resolve_server_entry — pure function
# ----------------------------------------------------------------------


class TestResolveServerEntry:
    def test_returns_dict_with_command(self) -> None:
        entry = resolve_server_entry()
        assert "command" in entry
        # Either the exe path OR sys.executable + args fallback
        if "args" in entry:
            assert entry["args"] == ["-m", "modelrisk_mcp"]


# ----------------------------------------------------------------------
# install — happy paths
# ----------------------------------------------------------------------


class TestInstallHappyPath:
    def test_creates_config_when_missing(self, tmp_path: Path) -> None:
        target = _make_target(tmp_path)
        assert not target.config_path.exists()

        results = install(
            clients=[target], server_entry={"command": "demo"},
        )
        assert len(results) == 1
        assert results[0].action == "added"
        assert results[0].backup_path is None  # nothing to back up

        # The file was created with our entry.
        data = json.loads(target.config_path.read_text())
        assert data == {"mcpServers": {"modelrisk": {"command": "demo"}}}

    def test_merges_into_existing_config(self, tmp_path: Path) -> None:
        """Pre-existing other servers must be preserved."""
        target = _make_target(tmp_path)
        target.config_path.write_text(json.dumps({
            "mcpServers": {
                "filesystem": {"command": "npx", "args": ["@mcp/filesystem"]},
            },
            "preferences": {"theme": "dark"},
        }))

        results = install(
            clients=[target], server_entry={"command": "demo"},
        )
        assert results[0].action == "added"
        assert results[0].backup_path is not None
        assert results[0].backup_path.exists()

        data = json.loads(target.config_path.read_text())
        # The other server stays untouched.
        assert data["mcpServers"]["filesystem"] == {
            "command": "npx", "args": ["@mcp/filesystem"],
        }
        # Our entry was added.
        assert data["mcpServers"]["modelrisk"] == {"command": "demo"}
        # Non-mcpServers content (preferences) is untouched.
        assert data["preferences"] == {"theme": "dark"}

    def test_updates_existing_modelrisk_entry_with_force(
        self, tmp_path: Path,
    ) -> None:
        target = _make_target(tmp_path)
        target.config_path.write_text(json.dumps({
            "mcpServers": {
                "modelrisk": {"command": "/old/path/modelrisk-mcp"},
            }
        }))

        results = install(
            clients=[target],
            server_entry={"command": "/new/path/modelrisk-mcp"},
            force=True,
        )
        assert results[0].action == "updated"

        data = json.loads(target.config_path.read_text())
        assert data["mcpServers"]["modelrisk"]["command"] == "/new/path/modelrisk-mcp"

    def test_skips_if_existing_entry_differs_without_force(
        self, tmp_path: Path,
    ) -> None:
        target = _make_target(tmp_path)
        target.config_path.write_text(json.dumps({
            "mcpServers": {
                "modelrisk": {"command": "/old/path/modelrisk-mcp"},
            }
        }))

        results = install(
            clients=[target],
            server_entry={"command": "/new/path/modelrisk-mcp"},
            force=False,
        )
        assert results[0].action == "skipped"
        # The existing entry must be untouched.
        data = json.loads(target.config_path.read_text())
        assert data["mcpServers"]["modelrisk"]["command"] == "/old/path/modelrisk-mcp"

    def test_no_op_when_existing_entry_matches(
        self, tmp_path: Path,
    ) -> None:
        """If the entry is already what we'd write, just `update` it
        idempotently — no skip, no error."""
        target = _make_target(tmp_path)
        target.config_path.write_text(json.dumps({
            "mcpServers": {"modelrisk": {"command": "demo"}},
        }))

        results = install(
            clients=[target], server_entry={"command": "demo"},
        )
        # entry matches → goes through the "added/updated" path (still
        # writes back, but doesn't change anything semantically).
        assert results[0].action == "updated"

    def test_custom_server_name(self, tmp_path: Path) -> None:
        target = _make_target(tmp_path)
        install(
            clients=[target],
            server_name="modelrisk-dev",
            server_entry={"command": "x"},
        )
        data = json.loads(target.config_path.read_text())
        assert "modelrisk-dev" in data["mcpServers"]
        assert "modelrisk" not in data["mcpServers"]


# ----------------------------------------------------------------------
# install — failure modes
# ----------------------------------------------------------------------


class TestInstallFailures:
    def test_no_clients_found_raises(self, tmp_path: Path) -> None:
        with pytest.raises(InstallError, match="No supported MCP clients"):
            install(clients=[])

    def test_malformed_existing_json_reports_error(
        self, tmp_path: Path,
    ) -> None:
        target = _make_target(tmp_path)
        target.config_path.write_text("{ this is not valid JSON")

        results = install(
            clients=[target], server_entry={"command": "demo"},
        )
        assert results[0].action == "error"
        assert "valid JSON" in results[0].message
        # Original (malformed) file must be untouched.
        assert "this is not valid JSON" in target.config_path.read_text()

    def test_non_object_servers_key_reports_error(
        self, tmp_path: Path,
    ) -> None:
        target = _make_target(tmp_path)
        target.config_path.write_text(json.dumps({"mcpServers": "oops"}))

        results = install(
            clients=[target], server_entry={"command": "demo"},
        )
        assert results[0].action == "error"
        assert "not an object" in results[0].message

    def test_empty_file_treated_as_empty_config(
        self, tmp_path: Path,
    ) -> None:
        target = _make_target(tmp_path)
        target.config_path.write_text("   \n  ")
        results = install(
            clients=[target], server_entry={"command": "demo"},
        )
        assert results[0].action == "added"


# ----------------------------------------------------------------------
# uninstall
# ----------------------------------------------------------------------


class TestUninstall:
    def test_removes_entry(self, tmp_path: Path) -> None:
        target = _make_target(tmp_path)
        target.config_path.write_text(json.dumps({
            "mcpServers": {
                "modelrisk": {"command": "demo"},
                "other": {"command": "x"},
            }
        }))

        results = uninstall(clients=[target])
        assert results[0].action == "removed"
        assert results[0].backup_path is not None

        data = json.loads(target.config_path.read_text())
        assert "modelrisk" not in data["mcpServers"]
        # Other servers preserved.
        assert "other" in data["mcpServers"]

    def test_idempotent_when_entry_missing(self, tmp_path: Path) -> None:
        target = _make_target(tmp_path)
        target.config_path.write_text(json.dumps({"mcpServers": {}}))

        results = uninstall(clients=[target])
        assert results[0].action == "skipped"
        assert "not present" in results[0].message

    def test_missing_file_skipped(self, tmp_path: Path) -> None:
        target = _make_target(tmp_path)
        # No file at all.
        results = uninstall(clients=[target])
        assert results[0].action == "skipped"
        assert "does not exist" in results[0].message


# ----------------------------------------------------------------------
# CLI integration — exercise the __main__ subcommand dispatch
# ----------------------------------------------------------------------


class TestCliDispatch:
    def test_install_subcommand_runs(
        self, tmp_path: Path, monkeypatch: pytest.MonkeyPatch
    ) -> None:
        """`modelrisk-mcp install` should reach the install machinery
        and exit 0 on a healthy run."""
        target = _make_target(tmp_path)

        def fake_discover() -> list[ClientTarget]:
            return [target]

        import modelrisk_mcp.install as install_mod
        monkeypatch.setattr(install_mod, "discover_clients", fake_discover)
        monkeypatch.setattr(
            install_mod, "resolve_server_entry",
            lambda: {"command": "demo"},
        )

        from modelrisk_mcp.__main__ import main as cli_main
        with pytest.raises(SystemExit) as exc_info:
            cli_main(["install"])
        assert exc_info.value.code == 0
        assert target.config_path.is_file()

    def test_serve_remains_default_no_subcommand(
        self, monkeypatch: pytest.MonkeyPatch
    ) -> None:
        """Backward compat: `modelrisk-mcp` with no args must still
        run the server (otherwise existing Claude Desktop configs
        break)."""
        from modelrisk_mcp import __main__ as main_mod

        called: dict[str, bool] = {}

        def fake_stdio() -> None:
            called["stdio"] = True

        monkeypatch.setattr(main_mod, "_run_stdio", fake_stdio)
        main_mod.main([])
        assert called.get("stdio") is True

    def test_uninstall_subcommand_runs(
        self, tmp_path: Path, monkeypatch: pytest.MonkeyPatch
    ) -> None:
        target = _make_target(tmp_path)
        target.config_path.write_text(json.dumps({
            "mcpServers": {"modelrisk": {"command": "demo"}}
        }))

        import modelrisk_mcp.install as install_mod
        monkeypatch.setattr(
            install_mod, "discover_clients", lambda: [target]
        )

        from modelrisk_mcp.__main__ import main as cli_main
        with pytest.raises(SystemExit) as exc_info:
            cli_main(["uninstall"])
        assert exc_info.value.code == 0
        data = json.loads(target.config_path.read_text())
        assert "modelrisk" not in data["mcpServers"]

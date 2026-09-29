"""Which MCP clients `install` / `status` find, and where each keeps its settings.

Every test builds a fake machine under tmp_path: a home folder, %APPDATA%,
%LOCALAPPDATA% with a Store package folder. Nothing touches the real
Claude, Cursor or VS Code settings of the machine running the tests.
"""

from __future__ import annotations

import json
from pathlib import Path

import pytest

import modelrisk_mcp.install as install_mod
from modelrisk_mcp.install import (
    InstallError,
    discover_clients,
    install,
    known_clients,
    select_clients,
    status,
)


@pytest.fixture
def machine(tmp_path: Path, monkeypatch: pytest.MonkeyPatch) -> Path:
    home, roaming, local = tmp_path / "home", tmp_path / "Roaming", tmp_path / "Local"
    for d in (home, roaming, local):
        d.mkdir()
    monkeypatch.setattr(install_mod.sys, "platform", "win32")
    monkeypatch.setattr(install_mod, "_home", lambda: home)
    monkeypatch.setattr(install_mod, "_claude_cli", lambda: None)
    monkeypatch.setenv("APPDATA", str(roaming))
    monkeypatch.setenv("LOCALAPPDATA", str(local))
    return tmp_path


def _store_folder(machine: Path) -> Path:
    folder = machine / "Local" / "Packages" / "Claude_pzs8sxrjxfjjc" / "LocalCache" / "Roaming" / "Claude"
    folder.mkdir(parents=True)
    return folder


def _ids(targets: list[install_mod.ClientTarget]) -> list[str]:
    return [t.id for t in targets]


class TestClaudeDesktop:
    def test_store_build_is_found_in_its_package_folder(self, machine: Path) -> None:
        """The Store build's settings are only in its package folder for
        anything outside the package (Excel, a terminal): %APPDATA%\\Claude
        does not exist there, and a file written to it is never read."""
        folder = _store_folder(machine)
        desktop = [t for t in discover_clients() if t.id == "claude-desktop"]
        assert len(desktop) == 1
        assert desktop[0].config_path == folder / "claude_desktop_config.json"
        assert not (machine / "Roaming" / "Claude").exists()

    def test_classic_install_is_found_in_appdata(self, machine: Path) -> None:
        (machine / "Roaming" / "Claude").mkdir()
        desktop = [t for t in discover_clients() if t.id == "claude-desktop"]
        assert desktop[0].config_path == machine / "Roaming" / "Claude" / "claude_desktop_config.json"

    def test_both_installs_are_two_targets(self, machine: Path) -> None:
        _store_folder(machine)
        (machine / "Roaming" / "Claude").mkdir()
        assert _ids(discover_clients())[:2] == ["claude-desktop", "claude-desktop-classic"]

    def test_neither_means_not_installed(self, machine: Path) -> None:
        assert "claude-desktop" not in _ids(discover_clients())
        assert "claude-desktop" in _ids(known_clients())   # still listed by status


class TestOtherClients:
    def test_each_is_found_by_its_settings_folder(self, machine: Path) -> None:
        home = machine / "home"
        for d in (home / ".cursor", machine / "Roaming" / "Code" / "User",
                  home / ".codeium" / "windsurf", home / ".gemini", home / ".lmstudio"):
            d.mkdir(parents=True)
        assert _ids(discover_clients()) == ["cursor", "vscode", "windsurf", "gemini-cli", "lm-studio"]

    def test_claude_code_needs_its_file_or_its_cli(self, machine: Path,
                                                    monkeypatch: pytest.MonkeyPatch) -> None:
        assert "claude-code" not in _ids(discover_clients())
        (machine / "home" / ".claude.json").write_text("{}", encoding="utf-8")
        assert "claude-code" in _ids(discover_clients())

    def test_vscode_entry_is_stdio_under_servers(self, machine: Path) -> None:
        (machine / "Roaming" / "Code" / "User").mkdir(parents=True)
        vscode = select_clients(["vscode"])
        results = install(clients=vscode, server_entry={"command": "C:/x/modelrisk-mcp.exe"})
        assert results[0].action == "added"
        data = json.loads((machine / "Roaming" / "Code" / "User" / "mcp.json").read_text(encoding="utf-8"))
        assert data == {"servers": {"modelrisk": {"type": "stdio", "command": "C:/x/modelrisk-mcp.exe"}}}

    def test_cursor_entry_is_the_plain_mcpservers_shape(self, machine: Path) -> None:
        (machine / "home" / ".cursor").mkdir()
        install(clients=select_clients(["cursor"]), server_entry={"command": "C:/x/modelrisk-mcp.exe"})
        data = json.loads((machine / "home" / ".cursor" / "mcp.json").read_text(encoding="utf-8"))
        assert data == {"mcpServers": {"modelrisk": {"command": "C:/x/modelrisk-mcp.exe"}}}


class TestSelectClients:
    def test_unknown_id_is_an_error(self, machine: Path) -> None:
        with pytest.raises(InstallError, match="Unknown client"):
            select_clients(["notepad"])

    def test_known_but_absent_is_an_error_not_a_skip(self, machine: Path) -> None:
        with pytest.raises(InstallError, match="not installed"):
            select_clients(["lm-studio"])


class TestStatus:
    def test_each_state(self, machine: Path) -> None:
        home = machine / "home"
        exe = machine / "bin" / "modelrisk-mcp.exe"
        exe.parent.mkdir()
        exe.write_bytes(b"")
        (home / ".cursor").mkdir()
        (home / ".cursor" / "mcp.json").write_text(
            json.dumps({"mcpServers": {"modelrisk": {"command": str(exe)}}}), encoding="utf-8")
        (home / ".gemini").mkdir()
        (home / ".gemini" / "settings.json").write_text(
            json.dumps({"mcpServers": {"modelrisk": {"command": str(machine / "gone.exe")}}}), encoding="utf-8")
        (home / ".codeium" / "windsurf").mkdir(parents=True)
        (home / ".codeium" / "windsurf" / "mcp_config.json").write_text("{,}", encoding="utf-8")
        (machine / "Roaming" / "Code" / "User").mkdir(parents=True)
        rows = {r["id"]: r for r in status()}
        assert rows["cursor"]["state"] == "added" and rows["cursor"]["command"] == str(exe)
        assert rows["gemini-cli"]["state"] == "broken"
        assert rows["windsurf"]["state"] == "unreadable"
        assert rows["vscode"]["state"] == "not_added"
        assert rows["lm-studio"] == {**rows["lm-studio"], "installed": False, "state": "not_installed"}

    def test_cli_json_is_the_contract(self, machine: Path, capsys: pytest.CaptureFixture[str],
                                      monkeypatch: pytest.MonkeyPatch) -> None:
        _store_folder(machine)
        monkeypatch.setattr(install_mod, "resolve_server_entry", lambda: {"command": "C:/x/modelrisk-mcp.exe"})
        from modelrisk_mcp.__main__ import main
        with pytest.raises(SystemExit) as exc:
            main(["status", "--json"])
        assert exc.value.code == 0
        out = json.loads(capsys.readouterr().out)
        assert out["version"] == 1 and out["server"] == "modelrisk"
        assert out["entry"] == {"command": "C:/x/modelrisk-mcp.exe"}
        desktop = next(c for c in out["clients"] if c["id"] == "claude-desktop")
        assert desktop["installed"] and desktop["state"] == "not_added"
        assert "LocalCache" in desktop["config_path"]


class TestCliClientFilter:
    def test_install_touches_only_the_named_client(self, machine: Path,
                                                    monkeypatch: pytest.MonkeyPatch) -> None:
        home = machine / "home"
        (home / ".cursor").mkdir()
        (home / ".gemini").mkdir()
        monkeypatch.setattr(install_mod, "resolve_server_entry", lambda: {"command": "C:/x/modelrisk-mcp.exe"})
        from modelrisk_mcp.__main__ import main
        with pytest.raises(SystemExit) as exc:
            main(["install", "--client", "cursor"])
        assert exc.value.code == 0
        assert (home / ".cursor" / "mcp.json").is_file()
        assert not (home / ".gemini" / "settings.json").exists()

    def test_uninstall_named_client(self, machine: Path, monkeypatch: pytest.MonkeyPatch) -> None:
        home = machine / "home"
        (home / ".cursor").mkdir()
        (home / ".cursor" / "mcp.json").write_text(
            json.dumps({"mcpServers": {"modelrisk": {"command": "x"}, "other": {"command": "y"}}}),
            encoding="utf-8")
        from modelrisk_mcp.__main__ import main
        with pytest.raises(SystemExit) as exc:
            main(["uninstall", "--client", "cursor"])
        assert exc.value.code == 0
        data = json.loads((home / ".cursor" / "mcp.json").read_text(encoding="utf-8"))
        assert data == {"mcpServers": {"other": {"command": "y"}}}

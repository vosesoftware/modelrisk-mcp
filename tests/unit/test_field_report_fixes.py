"""Regression tests for the 2026-07-20 field bug report (0.3.11):

1. Claude Code registration targets ~/.claude.json (NOT
   ~/.claude/settings.json) and prefers the `claude mcp add` CLI.
2. --read-only / MODELRISK_MCP_READ_ONLY is actually enforced.
3. Multi-instance Excel attach prefers the instance where ModelRisk
   answers the probe.
4. An MRService.dll lacking required exports fails with a version
   diagnosis, not a mid-call 'function not found'.
"""

from __future__ import annotations

from pathlib import Path
from unittest.mock import MagicMock, patch

import pytest

from modelrisk_mcp import install as install_mod
from modelrisk_mcp.bridge.excel import ExcelBridge, _instance_probe_live
from modelrisk_mcp.bridge.mrservice import MrServiceBridge
from modelrisk_mcp.config import Settings, read_only_active
from modelrisk_mcp.errors import ModelRiskNotLoadedError, ReadOnlyModeError

# ----------------------------------------------------------------------
# 1. Claude Code config target
# ----------------------------------------------------------------------


class TestClaudeCodeTarget:
    def test_config_path_is_claude_json_not_settings(self) -> None:
        p = install_mod._claude_code_config_path()
        assert p.name == ".claude.json"
        assert p.parent == Path.home()
        # The old, wrong location must be gone for good.
        assert "settings.json" not in str(p)

    def test_discover_keys_on_claude_json_or_cli(self, tmp_path: Path) -> None:
        with (
            patch.object(install_mod, "_claude_code_config_path",
                         return_value=tmp_path / ".claude.json"),
            patch.object(install_mod, "_claude_cli", return_value=None),
            patch.object(install_mod, "_claude_desktop_config_path",
                         return_value=tmp_path / "nowhere" / "cfg.json"),
        ):
            # Neither the file nor the CLI exists -> not detected.
            assert all(
                c.name != "Claude Code" for c in install_mod.discover_clients()
            )
            # File exists -> detected with the claude-cli strategy.
            (tmp_path / ".claude.json").write_text("{}", encoding="utf-8")
            clients = install_mod.discover_clients()
            code = next(c for c in clients if c.name == "Claude Code")
            assert code.strategy == "claude-cli"
            assert code.config_path.name == ".claude.json"

    def test_cli_preferred_when_present(self) -> None:
        entry = {"command": "C:/x/modelrisk-mcp.exe"}
        with (
            patch.object(install_mod, "_claude_cli", return_value="claude"),
            patch("subprocess.run") as run,
        ):
            run.return_value = MagicMock(returncode=0, stdout="", stderr="")
            res = install_mod._claude_cli_add("modelrisk", entry, force=False)
        assert res is not None and res.action == "added"
        cmd = run.call_args.args[0]
        assert cmd[:4] == ["claude", "mcp", "add", "modelrisk"]
        assert "--scope" in cmd and "user" in cmd
        assert "--" in cmd and "C:/x/modelrisk-mcp.exe" in cmd

    def test_json_fallback_writes_mcp_servers_in_claude_json(
        self, tmp_path: Path
    ) -> None:
        cfg = tmp_path / ".claude.json"
        cfg.write_text('{"other": 1}', encoding="utf-8")
        target = install_mod.ClientTarget(
            "Claude Code", cfg, strategy="claude-cli"
        )
        with patch.object(install_mod, "_claude_cli", return_value=None):
            res = install_mod.install(
                clients=[target], server_entry={"command": "x"}
            )
        assert res[0].action == "added"
        import json

        data = json.loads(cfg.read_text(encoding="utf-8"))
        assert data["mcpServers"]["modelrisk"] == {"command": "x"}
        assert data["other"] == 1  # untouched


# ----------------------------------------------------------------------
# 2. Read-only enforcement
# ----------------------------------------------------------------------


class TestReadOnly:
    def test_env_activates(self, monkeypatch: pytest.MonkeyPatch) -> None:
        monkeypatch.delenv("MODELRISK_MCP_READ_ONLY", raising=False)
        assert read_only_active() is False
        monkeypatch.setenv("MODELRISK_MCP_READ_ONLY", "1")
        assert read_only_active() is True

    def test_settings_flag_activates(
        self, monkeypatch: pytest.MonkeyPatch
    ) -> None:
        monkeypatch.delenv("MODELRISK_MCP_READ_ONLY", raising=False)
        assert read_only_active(Settings(read_only=True)) is True

    def test_write_cell_refuses(self, monkeypatch: pytest.MonkeyPatch) -> None:
        monkeypatch.setenv("MODELRISK_MCP_READ_ONLY", "true")
        with pytest.raises(ReadOnlyModeError, match="read-only"):
            ExcelBridge._ensure_writable("write cell Model!B6")

    def test_guard_passes_when_off(
        self, monkeypatch: pytest.MonkeyPatch
    ) -> None:
        monkeypatch.delenv("MODELRISK_MCP_READ_ONLY", raising=False)
        ExcelBridge._ensure_writable("write cell Model!B6")  # no raise


# ----------------------------------------------------------------------
# 3. Multi-instance attach
# ----------------------------------------------------------------------


def _fake_app(evaluate_result: object) -> MagicMock:
    app = MagicMock()
    if isinstance(evaluate_result, Exception):
        app.api.Evaluate.side_effect = evaluate_result
    else:
        app.api.Evaluate.return_value = evaluate_result
    return app


class TestMultiInstanceAttach:
    def test_probe_live_draw_vs_cverr(self) -> None:
        assert _instance_probe_live(_fake_app(7.0)) is True
        assert _instance_probe_live(_fake_app(0.0)) is True  # a valid draw
        assert _instance_probe_live(_fake_app(-2146826259)) is False  # #NAME?
        assert _instance_probe_live(_fake_app(RuntimeError("boom"))) is False

    def test_attach_prefers_functional_instance(self) -> None:
        bridge = ExcelBridge(auto_launch=False)
        dead = _fake_app(-2146826259)
        live = _fake_app(7.0)
        xw = MagicMock()
        xw.apps.__iter__ = lambda self: iter([dead, live])
        xw.apps.active = dead  # active points at the WRONG one
        bridge._xlwings = xw
        assert bridge._attach_active() is live

    def test_single_instance_uses_active(self) -> None:
        bridge = ExcelBridge(auto_launch=False)
        only = _fake_app(-2146826259)
        xw = MagicMock()
        xw.apps.__iter__ = lambda self: iter([only])
        xw.apps.active = only
        bridge._xlwings = xw
        assert bridge._attach_active() is only


# ----------------------------------------------------------------------
# 4. MRService export check
# ----------------------------------------------------------------------


class TestMrServiceExportCheck:
    def test_missing_ex2_export_diagnosed(self) -> None:
        class OldLib:
            """Mimics ctypes.CDLL for a 7.1.x-era DLL: Ex2 missing."""

            def __getattr__(self, name: str):
                if name == "MRLIB_SetOfflineActivationKeyEx2":
                    raise AttributeError(name)
                return MagicMock()

        with pytest.raises(ModelRiskNotLoadedError) as exc:
            MrServiceBridge._check_exports(OldLib(), r"C:\Tamara\MRService.dll")
        msg = str(exc.value)
        assert "too old" in msg
        assert "MRLIB_SetOfflineActivationKeyEx2" in msg
        assert "7.3.2.1" in msg

    def test_complete_dll_passes(self) -> None:
        class NewLib:
            def __getattr__(self, name: str):
                return MagicMock()

        MrServiceBridge._check_exports(NewLib(), "x.dll")  # no raise

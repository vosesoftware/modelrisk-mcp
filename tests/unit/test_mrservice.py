"""Tests for the new v0.3 MRService.dll bridge.

Without the actual DLL these tests stay focused on:
- Path discovery (env var + standard install paths)
- The activation flow's environment-variable contract
- Error message clarity when activation fails / DLL is missing

A separate integration test (gated on MRService.dll being installed)
exercises the full open-vmrs / read-samples flow.
"""

from __future__ import annotations

from pathlib import Path

import pytest

from modelrisk_mcp.bridge.mrservice import (
    MrServiceBridge,
    find_latest_vmrs,
    find_mrservice_dll,
)
from modelrisk_mcp.errors import ModelRiskNotLoadedError, SimulationFailedError


class TestFindDll:
    def test_returns_none_when_no_dll_found(
        self, monkeypatch: pytest.MonkeyPatch
    ) -> None:
        monkeypatch.delenv("MRSERVICE_DLL", raising=False)
        monkeypatch.setattr(
            "modelrisk_mcp.bridge.mrservice._DEFAULT_DLL_CANDIDATES",
            (r"C:\nonexistent\MRService.dll",),
        )
        assert find_mrservice_dll() is None

    def test_env_override_wins(
        self, monkeypatch: pytest.MonkeyPatch, tmp_path: Path
    ) -> None:
        fake = tmp_path / "MRService.dll"
        fake.write_bytes(b"")
        monkeypatch.setenv("MRSERVICE_DLL", str(fake))
        assert find_mrservice_dll() == str(fake)


class TestBridgeLifecycle:
    def test_missing_dll_raises_clear_error(
        self, monkeypatch: pytest.MonkeyPatch
    ) -> None:
        monkeypatch.delenv("MRSERVICE_DLL", raising=False)
        monkeypatch.setattr(
            "modelrisk_mcp.bridge.mrservice._DEFAULT_DLL_CANDIDATES",
            (r"C:\nonexistent\MRService.dll",),
        )
        bridge = MrServiceBridge()
        with pytest.raises(ModelRiskNotLoadedError) as exc:
            bridge.ensure_ready()
        msg = str(exc.value)
        assert "MRService.dll not found" in msg
        assert "MRSERVICE_DLL" in msg

    def test_missing_activation_key_raises(
        self, monkeypatch: pytest.MonkeyPatch, tmp_path: Path
    ) -> None:
        """If MRService.dll loads but no activation key env vars are
        set, activation should fail with an actionable error."""
        # We won't actually load a real DLL — mock the lib + skip _load.
        bridge = MrServiceBridge()
        bridge._lib = object()  # type: ignore[assignment]
        monkeypatch.delenv("MRSERVICE_ACTIVATION_KEY", raising=False)
        monkeypatch.delenv("MRSERVICE_ACTIVATION_KEY1", raising=False)
        monkeypatch.delenv("MRSERVICE_ACTIVATION_KEY2", raising=False)
        with pytest.raises(SimulationFailedError) as exc:
            bridge._activate()
        msg = str(exc.value)
        assert "MRSERVICE_ACTIVATION_KEY" in msg
        assert "activation" in msg.lower()

    def test_non_integer_key_rejected(
        self, monkeypatch: pytest.MonkeyPatch
    ) -> None:
        bridge = MrServiceBridge()
        bridge._lib = object()  # type: ignore[assignment]
        monkeypatch.setenv("MRSERVICE_ACTIVATION_KEY", "not-a-number")
        with pytest.raises(SimulationFailedError, match="not an integer"):
            bridge._activate()


class TestVmrsDiscovery:
    def test_returns_sibling_vmrs(self, tmp_path: Path) -> None:
        workbook = tmp_path / "model.xlsx"
        workbook.write_text("")
        vmrs = tmp_path / "model_1.vmrs"
        vmrs.write_bytes(b"")
        assert find_latest_vmrs(workbook) == str(vmrs)

    def test_returns_most_recent_when_multiple(self, tmp_path: Path) -> None:
        import time

        workbook = tmp_path / "model.xlsx"
        workbook.write_text("")
        old = tmp_path / "model_1.vmrs"
        old.write_bytes(b"")
        time.sleep(0.05)
        new = tmp_path / "model_2.vmrs"
        new.write_bytes(b"")
        assert find_latest_vmrs(workbook) == str(new)

    def test_returns_none_when_no_vmrs(self, tmp_path: Path) -> None:
        workbook = tmp_path / "model.xlsx"
        workbook.write_text("")
        assert find_latest_vmrs(workbook) is None

    def test_returns_none_when_workbook_missing(self, tmp_path: Path) -> None:
        assert find_latest_vmrs(tmp_path / "nope.xlsx") is None

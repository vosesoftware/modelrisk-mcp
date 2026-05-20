"""Tests for the add-in auto-loader.

Covers `ModelRiskBridge.ensure_modelrisk_active` end-to-end with a
mocked Excel that exposes a synthetic COMAddIns / AddIns collection.
"""

from __future__ import annotations

from typing import Any

import pytest

from modelrisk_mcp.bridge.modelrisk import ModelRiskBridge


class _FakeAddin:
    """Mimics an Office COMAddIn / AddIn object enough for our loader."""

    def __init__(
        self,
        *,
        description: str = "",
        progid: str = "",
        name: str = "",
        connected: bool = False,
        installed: bool = False,
    ) -> None:
        self.Description = description
        self.ProgID = progid
        self.Name = name
        self._connected = connected
        self._installed = installed
        self.Guid = ""

    @property
    def Connect(self) -> bool:  # noqa: N802
        return self._connected

    @Connect.setter
    def Connect(self, value: bool) -> None:  # noqa: N802
        self._connected = value

    @property
    def Installed(self) -> bool:  # noqa: N802
        return self._installed

    @Installed.setter
    def Installed(self, value: bool) -> None:  # noqa: N802
        self._installed = value

    @property
    def FullName(self) -> str:  # noqa: N802
        return ""


class _FakeExcel:
    """Exposes the minimal ExcelBridge surface the loader uses."""

    def __init__(
        self,
        com_addins: list[_FakeAddin],
        excel_addins: list[_FakeAddin],
    ) -> None:
        self._com = com_addins
        self._xll = excel_addins

    def list_com_addins(self) -> list[dict[str, Any]]:
        return [
            {
                "description": a.Description,
                "progid": a.ProgID,
                "guid": a.Guid,
                "connected": a._connected,
            }
            for a in self._com
        ]

    def list_excel_addins(self) -> list[dict[str, Any]]:
        return [
            {
                "name": a.Name,
                "installed": a._installed,
                "path": a.FullName,
            }
            for a in self._xll
        ]

    def enable_com_addin(self, predicate: Any) -> list[str]:
        flipped: list[str] = []
        for a in self._com:
            info = {
                "description": a.Description,
                "progid": a.ProgID,
                "guid": a.Guid,
                "connected": a._connected,
            }
            if not predicate(info) or info["connected"]:
                continue
            a.Connect = True
            flipped.append(info["description"] or info["progid"])
        return flipped

    def enable_excel_addin(self, predicate: Any) -> list[str]:
        flipped: list[str] = []
        for a in self._xll:
            info = {
                "name": a.Name,
                "installed": a._installed,
                "path": a.FullName,
            }
            if not predicate(info) or info["installed"]:
                continue
            a.Installed = True
            flipped.append(info["name"])
        return flipped


def _build_bridge(
    com_addins: list[_FakeAddin],
    excel_addins: list[_FakeAddin],
    *,
    dispatchable_after: bool,
    monkeypatch: pytest.MonkeyPatch,
) -> ModelRiskBridge:
    bridge = ModelRiskBridge(
        excel=_FakeExcel(com_addins, excel_addins),  # type: ignore[arg-type]
    )
    monkeypatch.setattr(bridge, "_try_dispatch", lambda: dispatchable_after)
    monkeypatch.setattr(
        bridge,
        "_try_dispatch_with_error",
        lambda: (dispatchable_after, None if dispatchable_after else "fake-error"),
    )
    return bridge


class TestEnsureModelriskActive:
    def test_enables_com_addin_named_modelrisk(
        self, monkeypatch: pytest.MonkeyPatch
    ) -> None:
        mr = _FakeAddin(
            description="ModelRisk Extended UI",
            progid="ModelRisk.ExtendedUI",
        )
        other = _FakeAddin(description="Acrobat PDF Toolbar")
        bridge = _build_bridge(
            com_addins=[mr, other],
            excel_addins=[],
            dispatchable_after=True,
            monkeypatch=monkeypatch,
        )
        diag = bridge.ensure_modelrisk_active()
        assert "ModelRisk Extended UI" in diag["com_addins_enabled"]
        assert mr._connected is True
        assert other._connected is False
        assert diag["modelrisk_dispatchable"] is True

    def test_enables_excel_xll_named_modelrisk(
        self, monkeypatch: pytest.MonkeyPatch
    ) -> None:
        mr = _FakeAddin(name="ModelRisk.xll")
        other = _FakeAddin(name="Analysis ToolPak")
        bridge = _build_bridge(
            com_addins=[],
            excel_addins=[mr, other],
            dispatchable_after=True,
            monkeypatch=monkeypatch,
        )
        diag = bridge.ensure_modelrisk_active()
        assert "ModelRisk.xll" in diag["excel_addins_enabled"]
        assert mr._installed is True
        assert other._installed is False

    def test_matches_vose_prefix_too(
        self, monkeypatch: pytest.MonkeyPatch
    ) -> None:
        vose = _FakeAddin(description="Vose Software COM Add-in")
        bridge = _build_bridge(
            com_addins=[vose],
            excel_addins=[],
            dispatchable_after=True,
            monkeypatch=monkeypatch,
        )
        diag = bridge.ensure_modelrisk_active()
        assert "Vose Software COM Add-in" in diag["com_addins_enabled"]

    def test_no_op_when_already_connected(
        self, monkeypatch: pytest.MonkeyPatch
    ) -> None:
        mr = _FakeAddin(description="ModelRisk", connected=True)
        bridge = _build_bridge(
            com_addins=[mr],
            excel_addins=[],
            dispatchable_after=True,
            monkeypatch=monkeypatch,
        )
        diag = bridge.ensure_modelrisk_active()
        # com_addins_enabled lists what we *flipped*, not what's on.
        assert diag["com_addins_enabled"] == []
        assert mr._connected is True

    def test_reports_unreachable_when_nothing_to_enable(
        self, monkeypatch: pytest.MonkeyPatch
    ) -> None:
        bridge = _build_bridge(
            com_addins=[_FakeAddin(description="Acrobat")],
            excel_addins=[_FakeAddin(name="Analysis ToolPak")],
            dispatchable_after=False,
            monkeypatch=monkeypatch,
        )
        diag = bridge.ensure_modelrisk_active()
        assert diag["com_addins_enabled"] == []
        assert diag["excel_addins_enabled"] == []
        assert diag["modelrisk_dispatchable"] is False
        # Diagnostic includes what we *did* see so the LLM can surface it.
        assert "Acrobat" in diag["com_addins_seen"]
        assert "Analysis ToolPak" in diag["excel_addins_seen"]


class TestIsModelriskLoadedAutoActivates:
    def test_calls_ensure_when_initial_dispatch_fails(
        self, monkeypatch: pytest.MonkeyPatch
    ) -> None:
        mr = _FakeAddin(description="ModelRisk Extended UI")
        bridge = ModelRiskBridge(
            excel=_FakeExcel([mr], []),  # type: ignore[arg-type]
        )
        dispatch_calls = {"count": 0}

        def _dispatch_factory() -> bool:
            dispatch_calls["count"] += 1
            # Simulate: first call fails, but once the add-in is on,
            # subsequent calls succeed.
            return mr._connected

        monkeypatch.setattr(bridge, "_try_dispatch", _dispatch_factory)
        assert bridge.is_modelrisk_loaded() is True
        assert dispatch_calls["count"] >= 2
        assert mr._connected is True

    def test_returns_false_when_no_modelrisk_addin_present(
        self, monkeypatch: pytest.MonkeyPatch
    ) -> None:
        bridge = ModelRiskBridge(
            excel=_FakeExcel([], []),  # type: ignore[arg-type]
        )
        monkeypatch.setattr(bridge, "_try_dispatch", lambda: False)
        assert bridge.is_modelrisk_loaded() is False


class TestStrategyDispatcher:
    """`_dispatch_via_first_working_strategy` falls through to the next
    strategy when the previous one returns ok=False."""

    def test_picks_first_working_strategy(
        self, monkeypatch: pytest.MonkeyPatch
    ) -> None:
        bridge = ModelRiskBridge(
            excel=_FakeExcel([], []),  # type: ignore[arg-type]
        )
        monkeypatch.setattr(
            bridge,
            "diagnose_dispatch_strategies",
            lambda: {
                "dispatch": {"ok": False, "error": "E_NOINTERFACE"},
                "dispatch_ex": {"ok": True, "error": None},
                "co_create": {"ok": False, "error": "not tried"},
                "via_comaddin": {"ok": False, "error": "not tried"},
            },
        )
        ok, strategy, err = bridge._dispatch_via_first_working_strategy()
        assert ok is True
        assert strategy == "dispatch_ex"
        assert err is None

    def test_falls_through_to_comaddin(
        self, monkeypatch: pytest.MonkeyPatch
    ) -> None:
        bridge = ModelRiskBridge(
            excel=_FakeExcel([], []),  # type: ignore[arg-type]
        )
        monkeypatch.setattr(
            bridge,
            "diagnose_dispatch_strategies",
            lambda: {
                "dispatch": {"ok": False, "error": "E_NOINTERFACE"},
                "dispatch_ex": {"ok": False, "error": "E_NOINTERFACE"},
                "co_create": {"ok": False, "error": "E_NOINTERFACE"},
                "via_comaddin": {"ok": True, "addin_name": "ModelRisk Ribbon"},
            },
        )
        ok, strategy, _ = bridge._dispatch_via_first_working_strategy()
        assert ok is True
        assert strategy == "via_comaddin"

    def test_reports_all_failures_when_nothing_works(
        self, monkeypatch: pytest.MonkeyPatch
    ) -> None:
        bridge = ModelRiskBridge(
            excel=_FakeExcel([], []),  # type: ignore[arg-type]
        )
        monkeypatch.setattr(
            bridge,
            "diagnose_dispatch_strategies",
            lambda: {
                "dispatch": {"ok": False, "error": "E_NOINTERFACE"},
                "dispatch_ex": {"ok": False, "error": "E_NOINTERFACE"},
                "co_create": {"ok": False, "error": "REGDB_E_CLASSNOTREG"},
                "via_comaddin": {"ok": False, "error": "no ModelRisk COMAddIn"},
            },
        )
        ok, strategy, err = bridge._dispatch_via_first_working_strategy()
        assert ok is False
        assert strategy is None
        # Summary mentions every attempt.
        assert "E_NOINTERFACE" in (err or "")
        assert "REGDB_E_CLASSNOTREG" in (err or "")

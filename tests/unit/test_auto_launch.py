"""Tests for auto-launching ModelRisk when no Excel is running.

The bridge now starts Vose's `modelrisk.exe` launcher (Excel comes up
with the add-in loaded natively) instead of erroring or relying on the
user to open Excel first. These tests drive the launch + attach loop
with a fake xlwings whose `apps.active` flips from None to an app once
"launched"."""

from __future__ import annotations

from typing import Any

import pytest

from modelrisk_mcp.bridge.excel import ExcelBridge
from modelrisk_mcp.errors import ExcelNotRunningError


class _FakeApps:
    def __init__(self, active: Any) -> None:
        self.active = active
        self.count = 1 if active is not None else 0


class _FakeXlwings:
    """Stand-in for the xlwings module. `apps.active` starts None and
    becomes a fake app object once `go_live()` is called (simulating
    Excel appearing after the launcher runs)."""

    def __init__(self) -> None:
        self._app = object()
        self.apps = _FakeApps(None)

    def go_live(self) -> None:
        self.apps = _FakeApps(self._app)


def _bridge(xw: _FakeXlwings, *, auto_launch: bool) -> ExcelBridge:
    b = ExcelBridge(auto_launch=auto_launch)
    b._xlwings = xw  # type: ignore[assignment]
    return b


class TestLaunchModelrisk:
    def test_returns_false_when_launcher_not_found(self, monkeypatch) -> None:
        xw = _FakeXlwings()
        b = _bridge(xw, auto_launch=True)
        monkeypatch.setattr(b, "_find_modelrisk_launcher", lambda: None)
        assert b.launch_modelrisk(timeout_s=1) is False

    def test_launches_and_attaches(self, monkeypatch) -> None:
        xw = _FakeXlwings()
        b = _bridge(xw, auto_launch=True)
        monkeypatch.setattr(
            b, "_find_modelrisk_launcher",
            lambda: r"C:\Program Files\Vose Software\ModelRisk\modelrisk.exe",
        )
        popen_calls: list[Any] = []

        def _fake_popen(args, **kw):
            popen_calls.append(args)
            xw.go_live()  # Excel "appears" right after launch
            return object()

        monkeypatch.setattr("subprocess.Popen", _fake_popen)
        assert b.launch_modelrisk(timeout_s=5) is True
        assert popen_calls and popen_calls[0][0].endswith("modelrisk.exe")
        assert b._app is not None  # attached

    def test_times_out_if_excel_never_appears(self, monkeypatch) -> None:
        xw = _FakeXlwings()  # stays None forever
        b = _bridge(xw, auto_launch=True)
        monkeypatch.setattr(
            b, "_find_modelrisk_launcher", lambda: r"C:\x\modelrisk.exe",
        )
        monkeypatch.setattr("subprocess.Popen", lambda *a, **k: object())
        assert b.launch_modelrisk(timeout_s=2) is False


class TestConnectAutoLaunch:
    def test_connect_auto_launches_when_no_excel(self, monkeypatch) -> None:
        xw = _FakeXlwings()
        b = _bridge(xw, auto_launch=True)
        called = {"n": 0}

        def _fake_launch(*, timeout_s: float = 45.0) -> bool:
            called["n"] += 1
            xw.go_live()
            b._app = xw.apps.active
            return True

        monkeypatch.setattr(b, "launch_modelrisk", _fake_launch)
        b.connect()
        assert called["n"] == 1
        assert b._app is not None

    def test_connect_raises_when_autolaunch_disabled(self) -> None:
        xw = _FakeXlwings()  # no Excel
        b = _bridge(xw, auto_launch=False)
        with pytest.raises(ExcelNotRunningError) as ei:
            b.connect()
        assert "auto-launch is disabled" in str(ei.value)

    def test_connect_attaches_without_launching_when_excel_present(
        self, monkeypatch
    ) -> None:
        xw = _FakeXlwings()
        xw.go_live()  # Excel already running
        b = _bridge(xw, auto_launch=True)
        launched = {"n": 0}
        monkeypatch.setattr(
            b, "launch_modelrisk",
            lambda **kw: launched.__setitem__("n", launched["n"] + 1) or True,
        )
        b.connect()
        assert launched["n"] == 0  # never needed to launch
        assert b._app is not None


class TestEnvToggle:
    def test_env_disables_auto_launch(self, monkeypatch) -> None:
        monkeypatch.setenv("MODELRISK_AUTO_LAUNCH", "0")
        assert ExcelBridge()._auto_launch is False

    def test_env_default_on(self, monkeypatch) -> None:
        monkeypatch.delenv("MODELRISK_AUTO_LAUNCH", raising=False)
        assert ExcelBridge()._auto_launch is True

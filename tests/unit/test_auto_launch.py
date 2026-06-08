"""Tests for auto-launching ModelRisk when no Excel is running.

The bridge starts an attachable Excel via ``xw.App(add_book=True)`` and
registers the ModelRisk XLL into it (rather than erroring or relying on
the user to open Excel first). A bare ``modelrisk.exe`` launch was tried
and rejected: it brings Excel up on the start screen with no workbook,
which is absent from the COM ROT and therefore unattachable. These tests
drive the launch + attach loop with a fake xlwings whose ``apps.active``
flips from None to an app once ``App(add_book=True)`` is called."""

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
    """Stand-in for the xlwings module. `apps.active` starts None; once
    `App(add_book=True)` is called it 'starts Excel' and apps.active
    becomes a fake app (an attachable Excel with a blank book)."""

    def __init__(self, *, app_raises: bool = False) -> None:
        self._app = object()
        self.apps = _FakeApps(None)
        self._app_raises = app_raises
        self.app_calls: list[dict[str, Any]] = []

    def App(self, **kwargs: Any) -> Any:  # noqa: N802 — mirrors xw.App
        self.app_calls.append(kwargs)
        if self._app_raises:
            raise RuntimeError("Excel failed to start")
        self.apps = _FakeApps(self._app)
        return self._app

    def go_live(self) -> None:
        self.apps = _FakeApps(self._app)


def _bridge(xw: _FakeXlwings, *, auto_launch: bool) -> ExcelBridge:
    b = ExcelBridge(auto_launch=auto_launch)
    b._xlwings = xw  # type: ignore[assignment]
    # The add-in-loading calls touch COM; stub them to no-ops for unit
    # tests (live XLL registration is exercised against real Excel).
    b.register_modelrisk_xlls = lambda: []  # type: ignore[method-assign]
    b.find_modelrisk_xll_paths = lambda: []  # type: ignore[method-assign]
    return b


class TestLaunchModelrisk:
    def test_starts_attachable_excel(self) -> None:
        xw = _FakeXlwings()
        b = _bridge(xw, auto_launch=True)
        assert b.launch_modelrisk() is True
        # Started Excel WITH a blank workbook so it's COM-attachable.
        assert xw.app_calls and xw.app_calls[0].get("add_book") is True
        assert b._app is not None

    def test_returns_false_when_excel_fails_to_start(self) -> None:
        xw = _FakeXlwings(app_raises=True)
        b = _bridge(xw, auto_launch=True)
        assert b.launch_modelrisk() is False

    def test_loads_addin_after_start(self) -> None:
        xw = _FakeXlwings()
        b = _bridge(xw, auto_launch=True)
        registered: list[str] = []
        b.register_modelrisk_xlls = lambda: []  # type: ignore[method-assign]
        b.find_modelrisk_xll_paths = (  # type: ignore[method-assign]
            lambda: [r"C:\Program Files\Vose Software\ModelRisk\ModelRisk.xll"]
        )
        b.register_xll = lambda p: registered.append(p) or True  # type: ignore[method-assign]
        assert b.launch_modelrisk() is True
        assert registered == [
            r"C:\Program Files\Vose Software\ModelRisk\ModelRisk.xll"
        ]


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

"""Tests for the ModelRisk add-in liveness probe + activation ladder
(bug #38). The whole point is to stop sims failing opaquely when the
add-in isn't loaded — so we verify the probe distinguishes live from
dead, and that the ladder escalates in the right order and raises an
actionable error when it can't recover."""

from __future__ import annotations

from typing import Any

import pytest

from modelrisk_mcp.bridge.modelrisk import ModelRiskBridge
from modelrisk_mcp.errors import ModelRiskNotFunctionalError

# A COM CVErr integer for #NAME? — what Evaluate returns for an
# unknown function (add-in not loaded).
NAME_ERR = -2146826259


class _FakeExcel:
    """Stand-in for ExcelBridge that scripts the add-in's liveness and
    records which activation steps were attempted."""

    def __init__(
        self,
        *,
        # sequence of probe outcomes Evaluate returns, consumed in order;
        # the last value repeats once exhausted
        evaluate_results: list[Any],
        installed_xlls: list[str] | None = None,
        offable_addins: list[str] | None = None,
        disk_xlls: list[str] | None = None,
        # which activation step (if any) flips the add-in to "live"
        live_after: str | None = None,
    ) -> None:
        self._evaluate_results = list(evaluate_results)
        self._installed_xlls = installed_xlls or []
        self._offable_addins = offable_addins or []
        self._disk_xlls = disk_xlls or []
        self._live_after = live_after
        self._live = False
        self.calls: list[str] = []
        self.connected = True

    def is_connected(self) -> bool:
        return self.connected

    def connect(self) -> None:
        self.connected = True

    def evaluate(self, expr: str) -> Any:
        if self._live:
            return 0.0  # live add-in: VoseNormal(0,1) static value
        if len(self._evaluate_results) > 1:
            return self._evaluate_results.pop(0)
        return self._evaluate_results[0]

    def register_modelrisk_xlls(self) -> list[str]:
        self.calls.append("register_installed")
        if self._live_after == "register" and self._installed_xlls:
            self._live = True
        return list(self._installed_xlls)

    def enable_excel_addin(self, predicate: Any) -> list[str]:
        self.calls.append("enable_offable")
        if self._live_after == "enable" and self._offable_addins:
            self._live = True
        return list(self._offable_addins)

    def find_modelrisk_xll_paths(self) -> list[str]:
        self.calls.append("find_on_disk")
        return list(self._disk_xlls)

    def register_xll(self, path: str) -> bool:
        self.calls.append(f"register_xll:{path}")
        if self._live_after == "disk":
            self._live = True
        return True


def _bridge(excel: _FakeExcel) -> ModelRiskBridge:
    b = ModelRiskBridge.__new__(ModelRiskBridge)
    b._excel = excel  # type: ignore[attr-defined]
    # is_modelrisk_loaded() (MRService) — stub to a constant
    b.is_modelrisk_loaded = lambda: True  # type: ignore[method-assign]
    return b


class TestProbe:
    def test_live_when_evaluate_returns_number(self) -> None:
        b = _bridge(_FakeExcel(evaluate_results=[0.0]))
        assert b.probe_addin_functional() is True

    def test_dead_on_name_error_code(self) -> None:
        b = _bridge(_FakeExcel(evaluate_results=[NAME_ERR]))
        assert b.probe_addin_functional() is False

    def test_dead_when_evaluate_raises(self) -> None:
        excel = _FakeExcel(evaluate_results=[0.0])
        excel.evaluate = lambda expr: (_ for _ in ()).throw(RuntimeError())  # type: ignore[method-assign]
        assert _bridge(excel).probe_addin_functional() is False

    def test_bool_is_not_treated_as_live(self) -> None:
        b = _bridge(_FakeExcel(evaluate_results=[True]))
        assert b.probe_addin_functional() is False


class TestEnsureLadder:
    def test_already_live_short_circuits(self) -> None:
        excel = _FakeExcel(evaluate_results=[0.0])
        h = _bridge(excel).ensure_modelrisk_functional()
        assert h.addin_functional is True
        assert excel.calls == []  # no activation attempted

    def test_recovers_by_registering_installed_xll(self) -> None:
        excel = _FakeExcel(
            evaluate_results=[NAME_ERR],
            installed_xlls=["ModelRisk64.xll"],
            live_after="register",
        )
        h = _bridge(excel).ensure_modelrisk_functional()
        assert h.addin_functional is True
        assert excel.calls == ["register_installed"]

    def test_recovers_by_enabling_offable_addin(self) -> None:
        excel = _FakeExcel(
            evaluate_results=[NAME_ERR],
            offable_addins=["ModelRisk"],
            live_after="enable",
        )
        h = _bridge(excel).ensure_modelrisk_functional()
        assert h.addin_functional is True
        assert excel.calls[:2] == ["register_installed", "enable_offable"]

    def test_recovers_by_registering_from_disk(self) -> None:
        excel = _FakeExcel(
            evaluate_results=[NAME_ERR],
            disk_xlls=[r"C:\Program Files\Vose Software\ModelRisk\ModelRisk64.xll"],
            live_after="disk",
        )
        h = _bridge(excel).ensure_modelrisk_functional()
        assert h.addin_functional is True
        assert "find_on_disk" in excel.calls
        assert any(c.startswith("register_xll:") for c in excel.calls)

    def test_raises_actionable_error_when_unrecoverable(self) -> None:
        excel = _FakeExcel(evaluate_results=[NAME_ERR])  # nothing recovers it
        with pytest.raises(ModelRiskNotFunctionalError) as ei:
            _bridge(excel).ensure_modelrisk_functional()
        msg = str(ei.value)
        assert "ModelRisk add-in isn't loaded" in msg
        assert "ribbon" in msg  # actionable guidance present

    def test_activate_false_raises_without_attempting(self) -> None:
        excel = _FakeExcel(evaluate_results=[NAME_ERR])
        with pytest.raises(ModelRiskNotFunctionalError):
            _bridge(excel).ensure_modelrisk_functional(activate=False)
        assert excel.calls == []


class TestHealthSnapshot:
    def test_health_is_probe_only_no_mutation(self) -> None:
        excel = _FakeExcel(evaluate_results=[NAME_ERR], installed_xlls=["x.xll"])
        h = _bridge(excel).health()
        assert h.addin_functional is False
        assert h.mrservice_ready is True
        assert excel.calls == []  # never tried to activate

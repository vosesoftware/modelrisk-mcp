"""Phase 4 simulation-tool tests with a mocked COM surface.

`FakeSimulationCom` records every call so we can assert:
- set_simulation_settings writes exactly the fields requested
- run_simulation calls settings + start in the right order
- a fixed seed implies use_fixed_seed=True (auto-flip)
- stop_simulation raises with the documented message
- get_simulation_status reflects in-process state during a run
"""

from __future__ import annotations

import threading
import time
from collections.abc import Iterator
from typing import Any

import pytest

from modelrisk_mcp.bridge.modelrisk import ModelRiskBridge
from modelrisk_mcp.bridge.simulation import SimulationController
from modelrisk_mcp.errors import SimulationNotAvailableError
from modelrisk_mcp.tools import reading, simulation


class FakeSimulationCom:
    """Records every call. Behaviour is deliberately deterministic so we
    can assert ordering and argument values."""

    def __init__(self) -> None:
        self.calls: list[tuple[str, tuple[Any, ...]]] = []
        self._block_event: threading.Event | None = None

    def __getattr__(self, name: str) -> Any:
        # Auto-record any set_* call.
        if name.startswith("set_"):

            def recorder(*args: Any) -> None:
                self.calls.append((name, args))

            return recorder
        raise AttributeError(name)

    def start(self) -> None:
        self.calls.append(("start", ()))
        if self._block_event is not None:
            # Block until the test releases us — used to verify
            # `get_simulation_status` returns "running" during a run.
            self._block_event.wait(timeout=5)


@pytest.fixture
def fake_com() -> FakeSimulationCom:
    return FakeSimulationCom()


@pytest.fixture
def bridge(
    fake_com: FakeSimulationCom,
    monkeypatch: pytest.MonkeyPatch,
) -> Iterator[ModelRiskBridge]:
    # ExcelBridge is irrelevant for the simulation tools — they go
    # through SimulationController only. Pass a typed-as-Any stub that
    # also pretends ModelRisk is "always loaded" so the simulation
    # tools' ensure_modelrisk_or_raise gate doesn't gatekeep us.
    class _NoExcel:
        def list_workbooks(self) -> list[Any]:
            return []

        def list_com_addins(self) -> list[dict[str, Any]]:
            return []

        def list_excel_addins(self) -> list[dict[str, Any]]:
            return []

        def enable_com_addin(self, predicate: Any) -> list[str]:
            return []

        def enable_excel_addin(self, predicate: Any) -> list[str]:
            return []

    bridge = ModelRiskBridge(
        excel=_NoExcel(),  # type: ignore[arg-type]
        simulation=SimulationController(com=fake_com),
    )
    # Real Dispatch would fail without Excel — short-circuit it for tests.
    monkeypatch.setattr(bridge, "_try_dispatch", lambda: True)
    reading.set_bridge_for_testing(bridge)
    yield bridge
    reading.set_bridge_for_testing(None)


# ----------------------------------------------------------------------
# set_simulation_settings
# ----------------------------------------------------------------------


class TestSetSimulationSettings:
    def test_only_requested_fields_written(
        self, fake_com: FakeSimulationCom, bridge: ModelRiskBridge
    ) -> None:
        response = simulation.set_simulation_settings(
            samples=5000, hide_progress_window=True
        )
        assert response.applied == {"samples": 5000, "hide_progress_window": True}
        # No spurious calls — only those two settings written.
        called_setters = [name for name, _ in fake_com.calls]
        assert called_setters == ["set_samples", "set_hide_progress_window"]

    def test_seed_implies_use_fixed_seed(
        self, fake_com: FakeSimulationCom, bridge: ModelRiskBridge
    ) -> None:
        response = simulation.set_simulation_settings(seed=42)
        assert response.applied["use_fixed_seed"] is True
        assert response.applied["seed"] == 42
        # use_fixed_seed=True must be set BEFORE the seed is written, so
        # ModelRisk treats the seed as authoritative.
        names = [n for n, _ in fake_com.calls]
        ufs_idx = names.index("set_use_fixed_seed")
        seed_idx = names.index("set_seed")
        assert ufs_idx < seed_idx

    def test_explicit_use_fixed_seed_false_with_seed_doesnt_override(
        self, fake_com: FakeSimulationCom, bridge: ModelRiskBridge
    ) -> None:
        """If the user explicitly says use_fixed_seed=False but supplies
        a seed, that's contradictory — we honour the explicit False
        rather than silently flipping it. The seed is still written so
        the user can flip later."""
        response = simulation.set_simulation_settings(
            use_fixed_seed=False, seed=42
        )
        # Only the explicit False got recorded for that flag.
        assert response.applied["use_fixed_seed"] is False
        assert response.applied["seed"] == 42
        # And the COM call sequence shows only one set_use_fixed_seed call.
        assert (
            sum(1 for n, _ in fake_com.calls if n == "set_use_fixed_seed") == 1
        )


# ----------------------------------------------------------------------
# run_simulation
# ----------------------------------------------------------------------


class TestRunSimulation:
    def test_calls_settings_then_start(
        self, fake_com: FakeSimulationCom, bridge: ModelRiskBridge
    ) -> None:
        response = simulation.run_simulation(iterations=1000)
        assert response.succeeded is True
        assert response.iterations_requested == 1000
        names = [n for n, _ in fake_com.calls]
        assert names == ["set_samples", "start"]
        # set_samples got the right value.
        samples_call = next(c for c in fake_com.calls if c[0] == "set_samples")
        assert samples_call[1] == (1000,)

    def test_seed_flips_use_fixed_seed_before_start(
        self, fake_com: FakeSimulationCom, bridge: ModelRiskBridge
    ) -> None:
        simulation.run_simulation(iterations=1000, seed=7)
        names = [n for n, _ in fake_com.calls]
        assert names == [
            "set_samples",
            "set_use_fixed_seed",
            "set_seed",
            "start",
        ]

    def test_no_args_just_starts(
        self, fake_com: FakeSimulationCom, bridge: ModelRiskBridge
    ) -> None:
        simulation.run_simulation()
        assert [n for n, _ in fake_com.calls] == ["start"]


# ----------------------------------------------------------------------
# stop_simulation
# ----------------------------------------------------------------------


class TestStopSimulation:
    def test_stop_raises_with_documented_message(
        self, bridge: ModelRiskBridge
    ) -> None:
        with pytest.raises(SimulationNotAvailableError) as exc:
            simulation.stop_simulation()
        message = str(exc.value)
        assert "Cancel" in message
        assert "progress dialog" in message


# ----------------------------------------------------------------------
# get_simulation_status
# ----------------------------------------------------------------------


class TestGetSimulationStatus:
    def test_idle_when_no_run_in_flight(
        self, bridge: ModelRiskBridge
    ) -> None:
        status = simulation.get_simulation_status()
        assert status.status == "idle"

    def test_running_during_run(
        self, fake_com: FakeSimulationCom, bridge: ModelRiskBridge
    ) -> None:
        block = threading.Event()
        fake_com._block_event = block

        result: dict[str, Any] = {}

        def run_in_background() -> None:
            simulation.run_simulation(iterations=100)
            result["done"] = True

        worker = threading.Thread(target=run_in_background)
        worker.start()
        # Spin briefly until the worker is inside `start()`.
        for _ in range(50):
            if any(c[0] == "start" for c in fake_com.calls):
                break
            time.sleep(0.01)
        assert simulation.get_simulation_status().status == "running"
        block.set()
        worker.join(timeout=5)
        assert result.get("done") is True
        assert simulation.get_simulation_status().status == "idle"

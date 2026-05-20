"""SimulationController — drives `ModelRiskSimulation` +
`ModelRiskSimulationSettings` over COM.

The COM dance is intentionally 3-step (Settings → StartSimulation →
Results) per spec §7.3. `_LiveSimulationCom` Dispatches both coclasses
lazily; tests inject a `FakeSimulationCom` to record method calls
without touching COM.

Status: `Application.StopSimulation()` and `Application.SimulationStatus`
are not yet exposed on the ModelRisk COM surface (see docs/com-surface.md
and spec §14 item 1). Until they are, `stop()` raises
`SimulationNotAvailableError` with the documented message, and `status()`
falls back to an in-process flag set during `run()`.
"""

from __future__ import annotations

import threading
from dataclasses import dataclass
from typing import Any, Protocol

from modelrisk_mcp.bridge.progids import (
    PROGID_SIMULATION,
    PROGID_SIMULATION_SETTINGS,
)
from modelrisk_mcp.errors import (
    SimulationFailedError,
    SimulationNotAvailableError,
)
from modelrisk_mcp.schemas.results import (
    SimulationRunResponse,
    SimulationSettingsRequest,
    SimulationSettingsResponse,
    SimulationStatus,
)

# A documented message reused by the not-yet-exposed stops.
_STOP_NOT_AVAILABLE_MSG: str = (
    "Simulation cancellation isn't exposed by the installed ModelRisk "
    "version. In current ModelRisk versions, press Cancel in the "
    "ModelRisk progress dialog."
)


class SimulationCom(Protocol):
    """Thin abstraction over the two simulation coclasses. Production
    uses pywin32 Dispatch; tests pass a fake to record calls."""

    def set_samples(self, n: int) -> None: ...
    def set_simulations(self, n: int) -> None: ...
    def set_use_fixed_seed(self, b: bool) -> None: ...
    def set_seed(self, index: int, value: float) -> None: ...
    def set_multiple_seed_type(self, t: int) -> None: ...
    def set_hide_progress_window(self, b: bool) -> None: ...
    def set_refresh_excel(self, b: bool) -> None: ...
    def set_refresh_rate(self, r: int) -> None: ...
    def set_stop_on_output_error(self, b: bool) -> None: ...
    def start(self) -> Any: ...


@dataclass
class _LiveSimulationCom:
    """Production SimulationCom — talks to the two ModelRisk coclasses."""

    _settings: Any = None
    _sim: Any = None

    def _get_settings(self) -> Any:
        if self._settings is None:
            self._settings = _dispatch(PROGID_SIMULATION_SETTINGS)
        return self._settings

    def _get_sim(self) -> Any:
        if self._sim is None:
            self._sim = _dispatch(PROGID_SIMULATION)
        return self._sim

    def set_samples(self, n: int) -> None:
        self._get_settings().Samples = n

    def set_simulations(self, n: int) -> None:
        self._get_settings().Simulations = n

    def set_use_fixed_seed(self, b: bool) -> None:
        self._get_settings().UseFixedSeed = b

    def set_seed(self, index: int, value: float) -> None:
        # Seed is a `propput, id(10) Seed(Seed_Index, newVal)` per the IDL.
        # pywin32 exposes this as a property accepting an index argument.
        self._get_settings().Seed[index] = value

    def set_multiple_seed_type(self, t: int) -> None:
        self._get_settings().MultipleSeedType = t

    def set_hide_progress_window(self, b: bool) -> None:
        self._get_settings().HideProgressWindow = b

    def set_refresh_excel(self, b: bool) -> None:
        self._get_settings().RefreshExcel = b

    def set_refresh_rate(self, r: int) -> None:
        self._get_settings().RefreshRate = r

    def set_stop_on_output_error(self, b: bool) -> None:
        self._get_settings().StopOnOutputError = b

    def start(self) -> Any:
        return self._get_sim().StartSimulation()


def _dispatch(progid: str) -> Any:
    try:
        import win32com.client as com
    except ImportError as exc:
        raise SimulationFailedError(
            "pywin32 is not installed; cannot reach ModelRisk's COM surface."
        ) from exc
    try:
        return com.Dispatch(progid)
    except Exception as exc:
        raise SimulationFailedError(
            f"Could not Dispatch {progid!r}. Is ModelRisk installed and "
            f"Excel running?"
        ) from exc


class SimulationController:
    """Coordinates simulation control. Exposes the four §7.3 operations:
    `apply_settings`, `run`, `stop`, `status`."""

    def __init__(self, com: SimulationCom | None = None) -> None:
        self._com: SimulationCom = com or _LiveSimulationCom()
        self._running = threading.Event()

    # ------------------------------------------------------------------
    # Settings
    # ------------------------------------------------------------------

    def apply_settings(
        self, request: SimulationSettingsRequest
    ) -> SimulationSettingsResponse:
        applied: dict[str, float | int | bool] = {}
        if request.samples is not None:
            self._com.set_samples(request.samples)
            applied["samples"] = request.samples
        if request.simulations is not None:
            self._com.set_simulations(request.simulations)
            applied["simulations"] = request.simulations
        if request.use_fixed_seed is not None:
            self._com.set_use_fixed_seed(request.use_fixed_seed)
            applied["use_fixed_seed"] = request.use_fixed_seed
        if request.seed is not None:
            # Setting a fixed seed only makes sense with UseFixedSeed=True.
            # Don't surprise the user — flip the flag for them if they
            # supplied a seed without explicitly setting use_fixed_seed.
            if request.use_fixed_seed is not False:
                self._com.set_use_fixed_seed(True)
                applied.setdefault("use_fixed_seed", True)
            self._com.set_seed(0, float(request.seed))
            applied["seed"] = request.seed
        if request.multiple_seed_type is not None:
            self._com.set_multiple_seed_type(request.multiple_seed_type)
            applied["multiple_seed_type"] = request.multiple_seed_type
        if request.hide_progress_window is not None:
            self._com.set_hide_progress_window(request.hide_progress_window)
            applied["hide_progress_window"] = request.hide_progress_window
        if request.refresh_excel is not None:
            self._com.set_refresh_excel(request.refresh_excel)
            applied["refresh_excel"] = request.refresh_excel
        if request.refresh_rate is not None:
            self._com.set_refresh_rate(request.refresh_rate)
            applied["refresh_rate"] = request.refresh_rate
        if request.stop_on_output_error is not None:
            self._com.set_stop_on_output_error(request.stop_on_output_error)
            applied["stop_on_output_error"] = request.stop_on_output_error
        return SimulationSettingsResponse(applied=applied)

    # ------------------------------------------------------------------
    # Run / stop / status
    # ------------------------------------------------------------------

    def run(
        self,
        iterations: int | None = None,
        seed: float | None = None,
    ) -> SimulationRunResponse:
        """Convenience method: optionally set samples + seed in the
        settings object, then call StartSimulation. Blocks until
        StartSimulation returns."""
        if iterations is not None:
            self._com.set_samples(iterations)
        if seed is not None:
            self._com.set_use_fixed_seed(True)
            self._com.set_seed(0, float(seed))
        self._running.set()
        try:
            self._com.start()
        except SimulationFailedError:
            raise
        except Exception as exc:
            raise SimulationFailedError(
                f"ModelRiskSimulation.StartSimulation() failed: {exc}"
            ) from exc
        finally:
            self._running.clear()
        return SimulationRunResponse(
            iterations_requested=iterations,
            seed_used=seed,
            succeeded=True,
        )

    def stop(self) -> None:
        """Cancellation isn't exposed on the current ModelRisk COM
        surface. We register the tool so the LLM can describe the limit
        but the call itself raises with the documented message."""
        raise SimulationNotAvailableError(_STOP_NOT_AVAILABLE_MSG)

    def status(self) -> SimulationStatus:
        if self._running.is_set():
            return SimulationStatus(
                status="running",
                note=(
                    "Coarse-grained: ModelRisk doesn't expose a SimulationStatus "
                    "property, so this only reflects in-process state. Once "
                    "StartSimulation returns, status flips to 'idle'."
                ),
            )
        return SimulationStatus(status="idle")

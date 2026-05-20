"""ResultsReader — extract simulation statistics from ModelRisk's COM
surface (the `ModelRiskSimulationResults` coclass and `ISimVariable`
proxies it returns).

See spec §8.4 and §7.4. Pearson/Spearman correlation and tornado
sensitivity are computed in Python (numpy) from `GetSamples()` arrays
because those endpoints aren't currently exposed on `ISimVariable` —
this is the [VOSE-INPUT-PENDING] item the developer is checking. If a
native path lands, the swap is a small refactor inside this module.
"""

from __future__ import annotations

import math
from collections.abc import Iterable
from dataclasses import dataclass
from typing import Any, Protocol

import numpy as np

from modelrisk_mcp.bridge.progids import PROGID_SIMULATION_RESULTS
from modelrisk_mcp.errors import SimulationFailedError
from modelrisk_mcp.schemas.results import (
    CorrelationMatrix,
    SensitivityEntry,
    SensitivityRanking,
    SimulationResult,
)

# ruff: noqa: N802 — COM method names are PascalCase by contract; matching
# them is required for duck-typing with `win32com.client.Dispatch` proxies.


class SimVariableLike(Protocol):
    """Minimal protocol for the methods we call on an ISimVariable. The
    production wrapper is a `win32com.client.Dispatch` proxy; tests can
    inject any object with the same shape."""

    def GetName(self) -> str: ...
    def GetMean(self) -> float: ...
    def GetVariance(self) -> float: ...
    def GetStDev(self) -> float: ...
    def GetSkewness(self) -> float: ...
    def GetKurtosis(self) -> float: ...
    def GetPercentile(self, p: float) -> float: ...
    def GetSamples(self) -> Any: ...


class ResultsCom(Protocol):
    """Abstracts the small slice of `ModelRiskSimulationResults` we use.
    Production: lazily Dispatches `PROGID_SIMULATION_RESULTS`. Tests
    inject a fake to avoid touching COM."""

    def sim_outputs(self) -> list[SimVariableLike]: ...
    def sim_inputs(self) -> list[SimVariableLike]: ...


@dataclass
class _LiveResultsCom:
    """Production ResultsCom — talks to ModelRisk over COM via pywin32.

    Lazily Dispatches on first use so unit tests can construct a
    ModelRiskBridge without Excel/ModelRisk being present.
    """

    _results: Any = None

    def _get(self) -> Any:
        if self._results is None:
            try:
                import win32com.client as com
            except ImportError as exc:
                raise SimulationFailedError(
                    "pywin32 is not installed; cannot read ModelRisk results."
                ) from exc
            try:
                self._results = com.Dispatch(PROGID_SIMULATION_RESULTS)
            except Exception as exc:
                raise SimulationFailedError(
                    "Could not Dispatch ModelRisk.ModelRiskSimulationResults. "
                    "Is ModelRisk installed and Excel running?"
                ) from exc
        return self._results

    def sim_outputs(self) -> list[SimVariableLike]:
        return self._collection_to_list(self._get().SimOutputs())

    def sim_inputs(self) -> list[SimVariableLike]:
        return self._collection_to_list(self._get().SimInputs())

    @staticmethod
    def _collection_to_list(collection: Any) -> list[SimVariableLike]:
        n = int(collection.Count)
        return [collection.Item(i) for i in range(1, n + 1)]


_DEFAULT_PERCENTILES: tuple[float, ...] = (
    0.05, 0.10, 0.25, 0.50, 0.75, 0.90, 0.95,
)


class ResultsReader:
    def __init__(self, com: ResultsCom | None = None) -> None:
        self._com: ResultsCom = com or _LiveResultsCom()

    # ------------------------------------------------------------------
    # Per-output statistics
    # ------------------------------------------------------------------

    def get_simulation_results(
        self,
        output_names: Iterable[str] | None = None,
        *,
        percentiles: tuple[float, ...] = _DEFAULT_PERCENTILES,
    ) -> list[SimulationResult]:
        wanted = set(output_names) if output_names else None
        results: list[SimulationResult] = []
        for var in self._com.sim_outputs():
            name = str(var.GetName())
            if wanted is not None and name not in wanted:
                continue
            samples = _as_float_array(var.GetSamples())
            results.append(
                SimulationResult(
                    output_name=name,
                    iterations=int(samples.size),
                    mean=_safe_float(var.GetMean()),
                    stdev=_safe_float(var.GetStDev()),
                    variance=_safe_optional_float(var.GetVariance()),
                    skewness=_safe_optional_float(var.GetSkewness()),
                    kurtosis=_safe_optional_float(var.GetKurtosis()),
                    min=float(np.min(samples)) if samples.size else 0.0,
                    max=float(np.max(samples)) if samples.size else 0.0,
                    percentiles={
                        p: _safe_float(var.GetPercentile(p)) for p in percentiles
                    },
                )
            )
        return results

    # ------------------------------------------------------------------
    # Correlation matrix (Python-computed; see spec §7.4)
    # ------------------------------------------------------------------

    def get_correlation_matrix(
        self, names: Iterable[str] | None = None
    ) -> CorrelationMatrix:
        variables = self._collect_named_variables(names)
        if not variables:
            return CorrelationMatrix()
        sample_matrix = self._stack_samples(variables)
        ordered_names = [name for name, _ in variables]
        pearson = _matrix_correlation(sample_matrix, kind="pearson")
        spearman = _matrix_correlation(sample_matrix, kind="spearman")
        return CorrelationMatrix(
            names=ordered_names,
            pearson=_matrix_to_optional_list(pearson),
            spearman=_matrix_to_optional_list(spearman),
            iterations=int(sample_matrix.shape[1]) if sample_matrix.size else 0,
        )

    # ------------------------------------------------------------------
    # Tornado (Python-computed; spec §7.4)
    # ------------------------------------------------------------------

    def get_sensitivity_ranking(self, output_name: str) -> SensitivityRanking:
        outputs = {str(v.GetName()): v for v in self._com.sim_outputs()}
        if output_name not in outputs:
            raise SimulationFailedError(
                f"Output {output_name!r} not found in simulation results."
            )
        output_samples = _as_float_array(outputs[output_name].GetSamples())
        n = output_samples.size
        entries: list[SensitivityEntry] = []
        for var in self._com.sim_inputs():
            in_name = str(var.GetName())
            in_samples = _as_float_array(var.GetSamples())
            if in_samples.size != n or n < 2:
                continue
            spearman = _spearman_pair(in_samples, output_samples)
            beta = _standardised_regression_coef(in_samples, output_samples)
            entries.append(
                SensitivityEntry(
                    input_name=in_name,
                    correlation=spearman,
                    regression_coefficient=beta,
                )
            )
        entries.sort(key=lambda e: abs(e.correlation), reverse=True)
        return SensitivityRanking(
            output_name=output_name, entries=entries, iterations=n
        )

    # ------------------------------------------------------------------
    # Internals
    # ------------------------------------------------------------------

    def _collect_named_variables(
        self, names: Iterable[str] | None
    ) -> list[tuple[str, SimVariableLike]]:
        wanted = set(names) if names else None
        variables: list[tuple[str, SimVariableLike]] = []
        for var in self._com.sim_inputs():
            name = str(var.GetName())
            if wanted is None or name in wanted:
                variables.append((name, var))
        for var in self._com.sim_outputs():
            name = str(var.GetName())
            if wanted is None or name in wanted:
                variables.append((name, var))
        return variables

    @staticmethod
    def _stack_samples(
        variables: list[tuple[str, SimVariableLike]],
    ) -> np.ndarray:
        if not variables:
            return np.zeros((0, 0))
        arrays = [_as_float_array(v.GetSamples()) for _, v in variables]
        n = min(a.size for a in arrays)
        if n == 0:
            return np.zeros((len(variables), 0))
        truncated = np.stack([a[:n] for a in arrays])
        return truncated


# ----------------------------------------------------------------------
# Helpers
# ----------------------------------------------------------------------


def _as_float_array(samples: Any) -> np.ndarray:
    """Normalise an `ISimVariable.GetSamples()` return value to a 1D
    numpy array of floats. ModelRisk may return a VARIANT-wrapped tuple
    of tuples (one row) — we flatten."""
    if samples is None:
        return np.array([], dtype=float)
    if hasattr(samples, "__iter__") and not isinstance(samples, (str, bytes)):
        try:
            return np.asarray(samples, dtype=float).flatten()
        except (TypeError, ValueError):
            # Fall through: maybe nested tuples — flatten manually.
            flat: list[float] = []
            for v in samples:
                if hasattr(v, "__iter__") and not isinstance(v, (str, bytes)):
                    for inner in v:
                        flat.append(float(inner))
                else:
                    flat.append(float(v))
            return np.asarray(flat, dtype=float)
    return np.asarray([samples], dtype=float)


def _safe_float(value: Any) -> float:
    if value is None:
        return float("nan")
    return float(value)


def _safe_optional_float(value: Any) -> float | None:
    if value is None:
        return None
    val = float(value)
    return val if math.isfinite(val) else None


def _matrix_correlation(
    samples_matrix: np.ndarray, *, kind: str
) -> np.ndarray:
    """Compute the k-by-k correlation matrix for a (k, n) sample matrix."""
    if samples_matrix.size == 0 or samples_matrix.shape[1] < 2:
        return np.full(
            (samples_matrix.shape[0], samples_matrix.shape[0]), np.nan
        )
    if kind == "spearman":
        ranked = np.apply_along_axis(_rankdata, 1, samples_matrix)
        return np.asarray(np.corrcoef(ranked))
    return np.asarray(np.corrcoef(samples_matrix))


def _matrix_to_optional_list(matrix: np.ndarray) -> list[list[float | None]]:
    out: list[list[float | None]] = []
    for row in matrix:
        out.append(
            [None if not math.isfinite(float(v)) else float(v) for v in row]
        )
    return out


def _rankdata(values: np.ndarray) -> np.ndarray:
    """Average-rank tie-breaking, matching scipy.stats.rankdata."""
    order = np.argsort(values, kind="mergesort")
    ranks = np.empty_like(order, dtype=float)
    ranks[order] = np.arange(1, len(values) + 1, dtype=float)
    # Average tied ranks
    sorted_values = values[order]
    i = 0
    while i < len(sorted_values):
        j = i + 1
        while j < len(sorted_values) and sorted_values[j] == sorted_values[i]:
            j += 1
        if j - i > 1:
            avg = (ranks[order[i:j]]).mean()
            ranks[order[i:j]] = avg
        i = j
    return ranks


def _spearman_pair(x: np.ndarray, y: np.ndarray) -> float:
    if x.size < 2:
        return 0.0
    rx = _rankdata(x)
    ry = _rankdata(y)
    if np.std(rx) == 0 or np.std(ry) == 0:
        return 0.0
    coef = float(np.corrcoef(rx, ry)[0, 1])
    return coef if math.isfinite(coef) else 0.0


def _standardised_regression_coef(x: np.ndarray, y: np.ndarray) -> float | None:
    if x.size < 2:
        return None
    sx = float(np.std(x, ddof=1))
    sy = float(np.std(y, ddof=1))
    if sx == 0 or sy == 0:
        return None
    pearson = float(np.corrcoef(x, y)[0, 1])
    if not math.isfinite(pearson):
        return None
    return pearson  # standardised beta = correlation when there's a single regressor

"""ResultsReader — pulls simulation statistics out of `.vmrs` files via
MRService.dll.

Architecture (v0.3, post-pivot from the ATL-COM approach):
- The .vmrs file is the source of truth. Excel + the XLL produce it
  when a simulation completes; we never try to drive ModelRisk's
  in-Excel COM surface again.
- `MrServiceBridge` (bridge/mrservice.py) gives us a clean ctypes
  surface for opening `.vmrs` files, reading samples, computing
  statistics, computing percentiles.
- Variable names → IDs go through `MRLIB_GetModelVarID`. The caller
  supplies which names to look up (typically pulled from Excel by
  `list_modelrisk_inputs` / `_outputs` first).
- Correlation matrix and tornado sensitivity are computed in Python
  (numpy) from `GetSamples` arrays — same as the COM path; that part
  doesn't change.
"""

from __future__ import annotations

import math
from collections.abc import Iterable

import numpy as np

from modelrisk_mcp.bridge.mrservice import (
    MrServiceBridge,
    VmrsHandle,
    find_latest_vmrs,
)
from modelrisk_mcp.errors import SimulationFailedError
from modelrisk_mcp.schemas.results import (
    CorrelationMatrix,
    SensitivityEntry,
    SensitivityRanking,
    SimulationResult,
)

_DEFAULT_PERCENTILES: tuple[float, ...] = (
    0.05, 0.10, 0.25, 0.50, 0.75, 0.90, 0.95,
)


# ----------------------------------------------------------------------
# Variable enumeration entry
# ----------------------------------------------------------------------


class VmrsVariableEntry:
    """One variable in a `.vmrs` confirmed against the open workbook.

    `kind` is "input" / "output" / "unknown" depending on which Excel
    list the name was found in. The `.vmrs` itself doesn't store these
    semantics — they come from the VoseInput / VoseOutput wrappers in
    the workbook."""

    __slots__ = ("iterations", "kind", "name", "var_id")

    def __init__(
        self, name: str, var_id: int, kind: str, iterations: int
    ) -> None:
        self.name = name
        self.var_id = var_id
        self.kind = kind
        self.iterations = iterations

    def to_dict(self) -> dict[str, str | int]:
        return {
            "name": self.name,
            "var_id": self.var_id,
            "kind": self.kind,
            "iterations": self.iterations,
        }


class ResultsReader:
    """Opens the active workbook's .vmrs file lazily and exposes the
    statistics surface the §7.1 reading tools need.

    `set_active_vmrs(path)` lets callers force a specific file (useful
    for the read_vmrs MCP tool). When not set, `_resolve_vmrs(workbook
    _path)` searches for a sibling `.vmrs` next to the workbook.
    """

    def __init__(
        self,
        mrservice: MrServiceBridge | None = None,
    ) -> None:
        self._mrservice = mrservice or MrServiceBridge()
        self._forced_vmrs: str | None = None

    def set_active_vmrs(self, path: str | None) -> None:
        self._forced_vmrs = path

    # ------------------------------------------------------------------
    # Public API — mirrors the old ResultsCom-based ResultsReader so the
    # tool layer doesn't change shape.
    # ------------------------------------------------------------------

    def get_simulation_results(
        self,
        workbook_path: str | None = None,
        output_names: Iterable[str] | None = None,
        *,
        percentiles: tuple[float, ...] = _DEFAULT_PERCENTILES,
    ) -> list[SimulationResult]:
        vmrs = self._resolve_vmrs(workbook_path)
        if vmrs is None:
            raise SimulationFailedError(
                "No .vmrs file found. Run a simulation in Excel first, "
                "then ask again. To target a specific file, call "
                "set_active_vmrs / use the read_vmrs tool."
            )
        wanted = list(output_names) if output_names else []
        results: list[SimulationResult] = []
        with self._mrservice.open_vmrs(vmrs) as handle:
            if not wanted:
                # No names supplied — enumerate every numeric var ID until
                # GetModelVarID stops finding matches. The .vmrs internally
                # has a count via Model_GetSimulations (1 typically) and
                # GetModelDataLength gives the iteration count, but the
                # variable count itself isn't directly exposed. We rely
                # on names being supplied by the caller in the normal flow;
                # this fallback just won't return anything sensible here.
                # Use the read_vmrs tool with explicit output_names for now.
                return []
            # Bug #23 (alpha.19): MRLIB_GetModelData appears to put the
            # handle into a state where subsequent MRLIB_GetModelVarID
            # calls return None — so we resolve EVERY name to a var_id
            # before pulling any samples. Without this, a fresh-from-
            # simulation .vmrs would silently lose all but the first
            # output (lookup → samples → lookup → fails → continue → ...).
            resolved: list[tuple[str, int]] = []
            for name in wanted:
                var_id = self._lookup_var_id(handle, name)
                if var_id is not None:
                    resolved.append((name, var_id))
            for name, var_id in resolved:
                samples = handle.get_samples(var_id)
                if not samples:
                    continue
                stats = handle.calc_statistics(samples)
                pcts = handle.calc_percentiles(samples, percentiles)
                results.append(
                    SimulationResult(
                        output_name=name,
                        iterations=stats.n,
                        mean=stats.mean,
                        stdev=stats.stdev,
                        variance=stats.variance,
                        skewness=stats.skewness,
                        kurtosis=stats.kurtosis,
                        min=stats.min,
                        max=stats.max,
                        percentiles=pcts,
                    )
                )
        return results

    def get_correlation_matrix(
        self,
        workbook_path: str | None = None,
        names: Iterable[str] | None = None,
    ) -> CorrelationMatrix:
        vmrs = self._resolve_vmrs(workbook_path)
        if vmrs is None or not names:
            return CorrelationMatrix()
        name_list = list(names)
        with self._mrservice.open_vmrs(vmrs) as handle:
            # Bug #23 (alpha.19): resolve every name to a var_id BEFORE
            # touching get_samples. See get_simulation_results for the
            # full backstory — MRLIB_GetModelData poisons subsequent
            # MRLIB_GetModelVarID lookups on the same handle.
            resolved: list[tuple[str, int]] = []
            for name in name_list:
                var_id = self._lookup_var_id(handle, name)
                if var_id is not None:
                    resolved.append((name, var_id))
            arrays: list[np.ndarray] = []
            ordered: list[str] = []
            for name, var_id in resolved:
                samples = handle.get_samples(var_id)
                if not samples:
                    continue
                arrays.append(np.asarray(samples, dtype=float))
                ordered.append(name)
            if not arrays:
                return CorrelationMatrix(names=name_list)
            n = min(a.size for a in arrays)
            matrix = np.stack([a[:n] for a in arrays])
            return CorrelationMatrix(
                names=ordered,
                pearson=_matrix_to_optional_list(_corrcoef(matrix)),
                spearman=_matrix_to_optional_list(_corrcoef(_rank_matrix(matrix))),
                iterations=int(n),
            )

    def list_variables(
        self,
        workbook_path: str | None,
        candidate_names: Iterable[tuple[str, str]],
    ) -> list[VmrsVariableEntry]:
        """Open the .vmrs and return the subset of `candidate_names` that
        actually have data. `candidate_names` is an iterable of
        `(name, kind)` pairs — kind is typically 'input' or 'output'.

        The SDK doesn't expose a name-enumeration call against a .vmrs,
        so we lookup-and-probe each candidate via MRLIB_GetModelVarID.
        That's why the caller must supply the candidates (usually pulled
        from VoseInput / VoseOutput cells in the open workbook)."""
        vmrs = self._resolve_vmrs(workbook_path)
        if vmrs is None:
            raise SimulationFailedError(
                "No .vmrs file found. Run a simulation in Excel first, "
                "then ask again. Or pin a specific file with set_active_vmrs."
            )
        out: list[VmrsVariableEntry] = []
        seen: set[str] = set()
        with self._mrservice.open_vmrs(vmrs) as handle:
            iterations = handle.iteration_count()
            for name, kind in candidate_names:
                if name in seen:
                    continue
                seen.add(name)
                var_id = self._lookup_var_id(handle, name)
                if var_id is None:
                    continue
                out.append(VmrsVariableEntry(name, var_id, kind, iterations))
        return out

    def get_samples(
        self,
        name: str,
        workbook_path: str | None = None,
        *,
        max_n: int = 10_000,
    ) -> tuple[float, ...]:
        """Return raw per-iteration samples for one variable. Filtered
        and errored samples are removed by MRService.dll. `max_n` caps
        the returned length for sanity over the MCP wire — a 1M-iteration
        sim would otherwise blow the JSON-RPC response budget."""
        vmrs = self._resolve_vmrs(workbook_path)
        if vmrs is None:
            raise SimulationFailedError(
                "No .vmrs file found. Run a simulation in Excel first, "
                "or pin a specific file with set_active_vmrs."
            )
        with self._mrservice.open_vmrs(vmrs) as handle:
            var_id = self._lookup_var_id(handle, name)
            if var_id is None:
                raise SimulationFailedError(
                    f"Variable {name!r} not found in {vmrs!r}. "
                    "Call list_vmrs_variables to see what's available."
                )
            return handle.get_samples(var_id, max_samples=max_n)

    def get_sensitivity_ranking(
        self,
        output_name: str,
        input_names: Iterable[str],
        workbook_path: str | None = None,
    ) -> SensitivityRanking:
        vmrs = self._resolve_vmrs(workbook_path)
        if vmrs is None:
            raise SimulationFailedError("No .vmrs available.")
        with self._mrservice.open_vmrs(vmrs) as handle:
            # Bug #23 (alpha.19): resolve EVERY name to a var_id before
            # touching get_samples. MRLIB_GetModelData appears to leave
            # the handle in a state where subsequent
            # MRLIB_GetModelVarID calls return None — first observed
            # on a fresh-from-simulation .vmrs where the output looked
            # up fine, its samples loaded, then every input lookup
            # returned None and the sensitivity ranking came back
            # empty. Resolving all IDs first sidesteps the issue.
            out_id = self._lookup_var_id(handle, output_name)
            if out_id is None:
                raise SimulationFailedError(
                    f"Output {output_name!r} not found in .vmrs."
                )
            resolved_inputs: list[tuple[str, int]] = []
            for in_name in input_names:
                in_id = self._lookup_var_id(handle, in_name)
                if in_id is not None:
                    resolved_inputs.append((in_name, in_id))

            out_samples = np.asarray(handle.get_samples(out_id), dtype=float)
            entries: list[SensitivityEntry] = []
            for in_name, in_id in resolved_inputs:
                in_samples = np.asarray(handle.get_samples(in_id), dtype=float)
                n = min(in_samples.size, out_samples.size)
                if n < 2:
                    continue
                spearman = _spearman_pair(in_samples[:n], out_samples[:n])
                beta = _standardised_regression_coef(
                    in_samples[:n], out_samples[:n]
                )
                entries.append(
                    SensitivityEntry(
                        input_name=in_name,
                        correlation=spearman,
                        regression_coefficient=beta,
                    )
                )
            entries.sort(key=lambda e: abs(e.correlation), reverse=True)
            return SensitivityRanking(
                output_name=output_name,
                entries=entries,
                iterations=int(out_samples.size),
            )

    # ------------------------------------------------------------------
    # Helpers
    # ------------------------------------------------------------------

    def _resolve_vmrs(self, workbook_path: str | None) -> str | None:
        if self._forced_vmrs:
            return self._forced_vmrs
        if workbook_path:
            return find_latest_vmrs(workbook_path)
        return None

    def _lookup_var_id(self, handle: VmrsHandle, name: str) -> int | None:
        """Resolve a variable name to its ID via MRLIB_GetModelVarID.
        Returns None if the function reports failure. The underlying
        call is wrapped in a wall-clock timeout (default 8 s) so a
        pathological name doesn't hang the entire MCP request — see
        `VmrsHandle.lookup_var_id`. On timeout, the underlying
        SimulationFailedError propagates with a clear message."""
        return handle.lookup_var_id(name)


# ----------------------------------------------------------------------
# numpy helpers — same as the old implementation
# ----------------------------------------------------------------------


def _corrcoef(matrix: np.ndarray) -> np.ndarray:
    """Wrap numpy.corrcoef so the result is always a 2-d array.

    Bug #28 (alpha.26): for a single-row input (1xN matrix —
    one variable), `np.corrcoef` returns a 0-d scalar of value 1.0
    (the correlation of one variable with itself). Downstream
    `_matrix_to_optional_list` then crashed with "iteration over
    a 0-d array" because it iterates over rows.

    Surfaces in real use: `get_correlation_matrix` called with a
    single name (or where only one of several names resolves) was
    crashing instead of returning a trivial [[1.0]] matrix."""
    if matrix.size == 0 or matrix.shape[1] < 2:
        return np.full((matrix.shape[0], matrix.shape[0]), np.nan)
    result = np.asarray(np.corrcoef(matrix))
    if result.ndim == 0:
        # Single-variable case — promote to a 1x1 matrix containing
        # the scalar (which numpy.corrcoef computed as 1.0 by
        # construction, but defer to the actual returned value rather
        # than hardcoding).
        return result.reshape(1, 1)
    return result


def _rank_matrix(matrix: np.ndarray) -> np.ndarray:
    return np.apply_along_axis(_rankdata, 1, matrix)


def _rankdata(values: np.ndarray) -> np.ndarray:
    order = np.argsort(values, kind="mergesort")
    ranks = np.empty_like(order, dtype=float)
    ranks[order] = np.arange(1, len(values) + 1, dtype=float)
    sorted_values = values[order]
    i = 0
    while i < len(sorted_values):
        j = i + 1
        while j < len(sorted_values) and sorted_values[j] == sorted_values[i]:
            j += 1
        if j - i > 1:
            avg = ranks[order[i:j]].mean()
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


def _standardised_regression_coef(
    x: np.ndarray, y: np.ndarray
) -> float | None:
    if x.size < 2:
        return None
    sx = float(np.std(x, ddof=1))
    sy = float(np.std(y, ddof=1))
    if sx == 0 or sy == 0:
        return None
    pearson = float(np.corrcoef(x, y)[0, 1])
    if not math.isfinite(pearson):
        return None
    return pearson


def _matrix_to_optional_list(matrix: np.ndarray) -> list[list[float | None]]:
    out: list[list[float | None]] = []
    for row in matrix:
        out.append(
            [None if not math.isfinite(float(v)) else float(v) for v in row]
        )
    return out


__all__ = ["ResultsReader"]

"""Direct ctypes wrapper for Vose Software's MRService.dll.

Verified architecture (2026-05-20 spike, docs/mrservice-spike.md):
- `MRLIB_SetOfflineActivationKey` activates the DLL per process.
  Single int64 key, env var `MRSERVICE_ACTIVATION_KEY`.
- `MRLIB_OpenSimulationModel(*.vmrs)` opens a saved simulation result
  file. `.xlsx` does NOT work — Excel + the XLL still has to run the
  simulation; we only read its output.
- `MRLIB_GetModelData(model, sim, var_id, buf, bufLen, samplesToRead,
  checkFilter, &nFilteredOut, &nErrorsOut)` reads the sample array.
- `MRLIB_CalcStatistics(data, n, _, &mean, &min, &max, &var, &cofV,
  &stdev, &skew, &kurt)` computes moments on a numpy/ctypes array.
- `MRLIB_CalcPercentilesArray(out, p_array, p_size, data, n, sorted,
  err)` computes arbitrary percentiles.

This module is the *only* thing that knows about MRService.dll. The
rest of the codebase talks to its public class `MrServiceBridge`.
"""

from __future__ import annotations

import ctypes
import os
from ctypes import (
    POINTER,
    byref,
    c_bool,
    c_double,
    c_int,
    c_int32,
    c_int64,
    c_longlong,
    c_wchar,
)
from dataclasses import dataclass
from pathlib import Path
from typing import Any

from modelrisk_mcp.bridge._keymat import decode_bundled_key
from modelrisk_mcp.errors import (
    ModelRiskNotLoadedError,
    SimulationFailedError,
)

# ---------------------------------------------------------------------------
# DLL discovery
# ---------------------------------------------------------------------------


_DLL_NAME: str = "MRService.dll"

# Bundled offline activation key. ModelRisk requires per-process activation
# of MRService.dll before it will open .vmrs files; rather than make every
# end user acquire and configure their own key, we ship one with the MCP
# server so the read path "just works". The plain integer is NOT in this
# file — it's reconstructed from obfuscated blobs in `_keymat` so the
# literal value isn't visible to `strings`, grep, or a curious wheel
# inspector. See `_keymat.py` for the rationale and the rotation script.
#
# `MRSERVICE_ACTIVATION_KEY` env var still takes precedence — useful for
# testing against a different key without rebuilding.

# Standard install paths in priority order. The env-var override is the
# canonical way for users to point at a non-standard install.
_DEFAULT_DLL_CANDIDATES: tuple[str, ...] = (
    r"C:\Program Files\Vose Software\ModelRisk\MRService.dll",
    r"C:\Program Files (x86)\Vose Software\ModelRisk\MRService.dll",
    # Repo-local SDK fallback for dev environments.
    r"C:\Users\timou\source\repos\ModelRisk\ModelRisk_Project\ModelRiskSDK\MRLibrary\_x64\MRService.dll",
    r"C:\Users\timou\source\repos\ModelRisk\ModelRisk_Project\ModelRiskSDK\MRLibrary\_x86\MRService.dll",
)


def find_mrservice_dll() -> str | None:
    """Locate MRService.dll. Honours the `MRSERVICE_DLL` env var first,
    then the standard ModelRisk install paths."""
    override = os.environ.get("MRSERVICE_DLL")
    if override and Path(override).is_file():
        return override
    for candidate in _DEFAULT_DLL_CANDIDATES:
        if Path(candidate).is_file():
            return candidate
    return None


# ---------------------------------------------------------------------------
# Result types
# ---------------------------------------------------------------------------


@dataclass(frozen=True)
class VmrsStatistics:
    """Output of `MRLIB_CalcStatistics`. Variance is the population
    variance; cofV is StDev/Mean."""

    n: int
    mean: float
    min: float
    max: float
    variance: float
    cof_v: float
    stdev: float
    skewness: float
    kurtosis: float


@dataclass(frozen=True)
class VmrsVariable:
    """One variable in a .vmrs file. `samples` is the raw per-iteration
    array (filtered + error samples removed)."""

    var_id: int
    name: str
    samples: tuple[float, ...]


# ---------------------------------------------------------------------------
# Bridge
# ---------------------------------------------------------------------------


class MrServiceBridge:
    """Owns a single MRService.dll handle.

    Lazy: the DLL is loaded and activated on first use. The activation
    key is read from `MRSERVICE_ACTIVATION_KEY` (or `MRSERVICE_ACTIVATION_KEY1/2`
    for the two-key variant). A bridge instance corresponds to one
    process's activation state; the DLL keeps it for the process
    lifetime.
    """

    def __init__(self, dll_path: str | None = None) -> None:
        self._dll_path = dll_path
        self._lib: ctypes.CDLL | None = None
        self._activated: bool = False

    # ----- lifecycle -----------------------------------------------------

    def ensure_ready(self) -> None:
        if self._lib is None:
            self._load()
        if not self._activated:
            self._activate()

    def _load(self) -> None:
        path = self._dll_path or find_mrservice_dll()
        if path is None:
            raise ModelRiskNotLoadedError(
                "MRService.dll not found. Set MRSERVICE_DLL to the full "
                "path of the DLL, or install ModelRisk."
            )
        # add_dll_directory makes License.dll resolvable.
        if hasattr(os, "add_dll_directory"):
            try:
                os.add_dll_directory(str(Path(path).parent))
            except OSError:
                pass
        try:
            self._lib = ctypes.cdll.LoadLibrary(path)
        except OSError as exc:
            raise ModelRiskNotLoadedError(
                f"Could not LoadLibrary({path!r}): {exc}"
            ) from exc
        self._configure_signatures(self._lib)
        self._dll_path = path

    @staticmethod
    def _configure_signatures(lib: ctypes.CDLL) -> None:
        # Activation
        lib.MRLIB_SetOfflineActivationKey.restype = c_bool
        lib.MRLIB_SetOfflineActivationKey.argtypes = [c_int64]
        lib.MRLIB_SetOfflineActivationKeyEx2.restype = c_bool
        lib.MRLIB_SetOfflineActivationKeyEx2.argtypes = [
            c_int64, c_int64,
            POINTER(c_int), POINTER(c_int), POINTER(c_int),
        ]
        # Model lifecycle
        lib.MRLIB_OpenSimulationModel.restype = c_bool
        lib.MRLIB_OpenSimulationModel.argtypes = [POINTER(c_wchar), POINTER(c_longlong)]
        lib.MRLIB_CloseSimulationModel.restype = c_bool
        lib.MRLIB_CloseSimulationModel.argtypes = [c_longlong]
        # Data access
        lib.MRLIB_GetModelDataLength.restype = c_int
        lib.MRLIB_GetModelDataLength.argtypes = [c_longlong, c_int]
        lib.MRLIB_GetModelData.restype = c_int
        lib.MRLIB_GetModelData.argtypes = [
            c_longlong, c_int, c_int, POINTER(c_double), c_int, c_int, c_int,
            POINTER(c_int), POINTER(c_int),
        ]
        # Statistics
        lib.MRLIB_CalcStatistics.restype = c_int
        lib.MRLIB_CalcStatistics.argtypes = [
            POINTER(c_double), c_int, POINTER(c_int),
            POINTER(c_double), POINTER(c_double), POINTER(c_double),
            POINTER(c_double), POINTER(c_double), POINTER(c_double),
            POINTER(c_double), POINTER(c_double),
        ]
        lib.MRLIB_CalcPercentilesArray.restype = c_int
        lib.MRLIB_CalcPercentilesArray.argtypes = [
            POINTER(c_double), POINTER(c_double), c_int,
            POINTER(c_double), c_int, c_int, POINTER(c_wchar),
        ]
        lib.MRLIB_CalculateRiskRatio.restype = c_int
        lib.MRLIB_CalculateRiskRatio.argtypes = [
            POINTER(c_double), POINTER(c_double), c_int, c_int,
        ]

    def _activate(self) -> None:
        assert self._lib is not None
        # Precedence:
        #   1. MRSERVICE_ACTIVATION_KEY (single int64) — env override
        #   2. MRSERVICE_ACTIVATION_KEY1/2 (split int64s) — env override
        #   3. _BUNDLED_ACTIVATION_KEY — ships with the MCP server so the
        #      read path is plug-and-play. Tests can suppress this by
        #      setting MRSERVICE_DISABLE_BUNDLED_KEY=1.
        single = os.environ.get("MRSERVICE_ACTIVATION_KEY")
        if single:
            try:
                key = c_int64(int(single))
            except ValueError as exc:
                raise SimulationFailedError(
                    f"MRSERVICE_ACTIVATION_KEY {single!r} is not an integer."
                ) from exc
            ok = bool(self._lib.MRLIB_SetOfflineActivationKey(key))
            if not ok:
                raise SimulationFailedError(
                    "MRLIB_SetOfflineActivationKey returned False. "
                    "The activation key was rejected."
                )
            self._activated = True
            return
        key1 = os.environ.get("MRSERVICE_ACTIVATION_KEY1")
        key2 = os.environ.get("MRSERVICE_ACTIVATION_KEY2")
        if key1 and key2:
            year, month, day = c_int(), c_int(), c_int()
            try:
                ok = bool(
                    self._lib.MRLIB_SetOfflineActivationKeyEx2(
                        c_int64(int(key1)), c_int64(int(key2)),
                        byref(year), byref(month), byref(day),
                    )
                )
            except ValueError as exc:
                raise SimulationFailedError(
                    "MRSERVICE_ACTIVATION_KEY1 / _2 must be integers."
                ) from exc
            if not ok:
                raise SimulationFailedError(
                    "MRLIB_SetOfflineActivationKeyEx2 returned False."
                )
            self._activated = True
            return
        if not os.environ.get("MRSERVICE_DISABLE_BUNDLED_KEY"):
            # Decode then immediately consume — don't keep the plain key
            # in a Python binding any longer than necessary.
            bundled = c_int64(decode_bundled_key())
            ok = bool(self._lib.MRLIB_SetOfflineActivationKey(bundled))
            del bundled
            if ok:
                self._activated = True
                return
            # Bundled key rejected — fall through to the clear error
            # rather than silently leaving the bridge unactivated.
            raise SimulationFailedError(
                "Bundled activation key was rejected by MRService.dll. "
                "Your installed ModelRisk SDK may be too new/old for this "
                "key. Set MRSERVICE_ACTIVATION_KEY to override, or "
                "report this to the modelrisk-mcp maintainers."
            )
        raise SimulationFailedError(
            "MRService.dll requires activation. Set MRSERVICE_ACTIVATION_KEY "
            "(a single int64) or both MRSERVICE_ACTIVATION_KEY1 and "
            "MRSERVICE_ACTIVATION_KEY2 (two int64s). MRSERVICE_DISABLE_"
            "BUNDLED_KEY is set, so the bundled key was not tried."
        )

    # ----- vmrs read API -------------------------------------------------

    def open_vmrs(self, path: str) -> VmrsHandle:
        """Open a .vmrs file. Returns a handle whose methods read the
        underlying simulation data; the handle must be closed (use it
        as a context manager)."""
        self.ensure_ready()
        assert self._lib is not None
        model_ptr = (c_longlong * 1)()
        ok = bool(self._lib.MRLIB_OpenSimulationModel(path, model_ptr))
        if not ok:
            raise SimulationFailedError(
                f"MRLIB_OpenSimulationModel({path!r}) returned False. "
                "The file may be missing, corrupted, or produced by an "
                "incompatible ModelRisk version."
            )
        return VmrsHandle(self._lib, model_ptr[0], path)


class VmrsHandle:
    """Wraps one MRLIB_OpenSimulationModel result. Closes the underlying
    model when used as a context manager or via `.close()`."""

    def __init__(self, lib: ctypes.CDLL, model_ptr: int, path: str) -> None:
        self._lib = lib
        self._model_ptr = model_ptr
        self._path = path
        self._closed = False

    def __enter__(self) -> VmrsHandle:
        return self

    def __exit__(self, *exc: Any) -> None:
        self.close()

    def close(self) -> None:
        if self._closed:
            return
        try:
            self._lib.MRLIB_CloseSimulationModel(c_longlong(self._model_ptr))
        finally:
            self._closed = True

    @property
    def path(self) -> str:
        return self._path

    @property
    def model_ptr(self) -> int:
        return self._model_ptr

    def iteration_count(self, sim_index: int = 0) -> int:
        """Number of iterations recorded for the given simulation."""
        n = int(
            self._lib.MRLIB_GetModelDataLength(c_longlong(self._model_ptr), c_int(sim_index))
        )
        return max(0, n)

    def get_samples(
        self,
        var_id: int,
        sim_index: int = 0,
        max_samples: int | None = None,
    ) -> tuple[float, ...]:
        """Return the per-iteration sample array for one variable.
        Filtered + errored samples are removed."""
        n_total = self.iteration_count(sim_index)
        if n_total == 0:
            return ()
        n = min(n_total, max_samples) if max_samples else n_total
        buf = (c_double * n)()
        n_filtered = c_int(0)
        n_errors = c_int(0)
        valid = int(
            self._lib.MRLIB_GetModelData(
                c_longlong(self._model_ptr),
                c_int(sim_index),
                c_int(var_id),
                buf,
                c_int(n),
                c_int(n),
                c_int(1),  # check filter + errors
                byref(n_filtered),
                byref(n_errors),
            )
        )
        valid = max(0, min(valid, n))
        return tuple(buf[i] for i in range(valid))

    def calc_statistics(self, samples: tuple[float, ...]) -> VmrsStatistics:
        n = len(samples)
        if n == 0:
            return VmrsStatistics(0, 0, 0, 0, 0, 0, 0, 0, 0)
        arr = (c_double * n)(*samples)
        flags = (c_int32 * n)()
        mean = c_double(0)
        v_min = c_double(0)
        v_max = c_double(0)
        variance = c_double(0)
        cof_v = c_double(0)
        stdev = c_double(0)
        skew = c_double(0)
        kurt = c_double(0)
        self._lib.MRLIB_CalcStatistics(
            arr, c_int(n), flags,
            byref(mean), byref(v_min), byref(v_max),
            byref(variance), byref(cof_v), byref(stdev),
            byref(skew), byref(kurt),
        )
        return VmrsStatistics(
            n=n,
            mean=mean.value,
            min=v_min.value,
            max=v_max.value,
            variance=variance.value,
            cof_v=cof_v.value,
            stdev=stdev.value,
            skewness=skew.value,
            kurtosis=kurt.value,
        )

    def calc_percentiles(
        self, samples: tuple[float, ...], percentiles: tuple[float, ...]
    ) -> dict[float, float]:
        """Compute arbitrary percentiles. `percentiles` is in [0, 1]."""
        n = len(samples)
        n_p = len(percentiles)
        if n == 0 or n_p == 0:
            return {}
        data = (c_double * n)(*samples)
        p_in = (c_double * n_p)(*percentiles)
        p_out = (c_double * n_p)()
        err_buf = (c_wchar * 64)()
        self._lib.MRLIB_CalcPercentilesArray(
            p_out, p_in, c_int(n_p),
            data, c_int(n),
            c_int(0),  # not pre-sorted
            err_buf,
        )
        return {percentiles[i]: p_out[i] for i in range(n_p)}


# ---------------------------------------------------------------------------
# .vmrs discovery
# ---------------------------------------------------------------------------


def find_latest_vmrs(workbook_path: str | Path) -> str | None:
    """ModelRisk writes a sibling `.vmrs` next to a workbook when a
    simulation completes (convention: `<book>_<n>.vmrs`). This locates
    the most recent one, falling back to any `.vmrs` in the same folder
    if naming doesn't match."""
    workbook = Path(workbook_path)
    if not workbook.is_file():
        return None
    folder = workbook.parent
    stem = workbook.stem
    candidates = sorted(
        folder.glob(f"{stem}*.vmrs"),
        key=lambda p: p.stat().st_mtime,
        reverse=True,
    )
    if not candidates:
        candidates = sorted(
            folder.glob("*.vmrs"),
            key=lambda p: p.stat().st_mtime,
            reverse=True,
        )
    return str(candidates[0]) if candidates else None


__all__ = [
    "MrServiceBridge",
    "VmrsHandle",
    "VmrsStatistics",
    "VmrsVariable",
    "find_latest_vmrs",
    "find_mrservice_dll",
]

"""Spike: can MRService.dll drive a ModelRisk workbook headlessly?

Phase B of the architectural pivot. We've confirmed MRService.dll
exposes a clean ctypes surface for reading `.vmrs` files. The open
question is whether `MRLIB_OpenSimulationModel(workbook_path)` can also
load a `.xlsx` directly — if yes, the whole simulation can run without
Excel being open. If no, we trigger simulation via the XLL inside
Excel and read the resulting `.vmrs` via MRService.dll.

Run from a Windows machine with ModelRisk installed:

    uv run python scripts/spike_mrservice.py path/to/test.xlsx

The script reports:
  - DLL load: OK / failed
  - Activation status: already activated / needs activation
  - Open .xlsx: success → full headless is viable
  - Open .vmrs (auto-located): success → read-only path is viable
  - Variable inspection on whichever model loaded

Output goes to stdout and to `docs/mrservice-spike.md`.
"""

from __future__ import annotations

import ctypes
import sys
from ctypes import POINTER, byref, c_bool, c_double, c_int, c_int64, c_longlong, c_wchar
from datetime import UTC, datetime
from pathlib import Path

DLL_CANDIDATES: tuple[str, ...] = (
    # Vose Software install paths (most common)
    r"C:\Program Files\Vose Software\ModelRisk\MRService.dll",
    r"C:\Program Files (x86)\Vose Software\ModelRisk\MRService.dll",
    # Repo-local SDK location
    r"C:\Users\timou\source\repos\ModelRisk\ModelRisk_Project\ModelRiskSDK\MRLibrary\_x64\MRService.dll",
    r"C:\Users\timou\source\repos\ModelRisk\ModelRisk_Project\ModelRiskSDK\MRLibrary\_x86\MRService.dll",
)

OUT_PATH = Path(__file__).resolve().parent.parent / "docs" / "mrservice-spike.md"


def _find_dll() -> str | None:
    for candidate in DLL_CANDIDATES:
        if Path(candidate).is_file():
            return candidate
    return None


def _load(path: str) -> ctypes.CDLL:
    # add_dll_directory makes the DLL's dependencies (License.dll etc.) resolvable.
    import os
    if hasattr(os, "add_dll_directory"):
        os.add_dll_directory(str(Path(path).parent))
    return ctypes.cdll.LoadLibrary(path)


def _setup_signatures(lib: ctypes.CDLL) -> None:
    lib.MRLIB_CreateSimulationModel.restype = c_bool
    lib.MRLIB_CreateSimulationModel.argtypes = [POINTER(c_wchar), POINTER(c_longlong)]
    lib.MRLIB_OpenSimulationModel.restype = c_bool
    lib.MRLIB_OpenSimulationModel.argtypes = [POINTER(c_wchar), POINTER(c_longlong)]
    lib.MRLIB_CloseSimulationModel.restype = c_bool
    lib.MRLIB_CloseSimulationModel.argtypes = [c_longlong]
    lib.MRLIB_GetModelDataLength.restype = c_int
    lib.MRLIB_GetModelDataLength.argtypes = [c_longlong, c_int]
    lib.MRLIB_GetModelData.restype = c_int
    lib.MRLIB_GetModelData.argtypes = [
        c_longlong, c_int, c_int, POINTER(c_double), c_int, c_int, c_int,
        POINTER(c_int), POINTER(c_int),
    ]
    lib.MRLIB_CalcStatistics.restype = c_int
    lib.MRLIB_CalcStatistics.argtypes = [
        POINTER(c_double), c_int, POINTER(c_int),
        POINTER(c_double), POINTER(c_double), POINTER(c_double),
        POINTER(c_double), POINTER(c_double), POINTER(c_double),
        POINTER(c_double), POINTER(c_double),
    ]
    # Licensing — may or may not be needed; we surface either outcome.
    lib.MRLIB_SetOfflineActivationKey.restype = c_bool
    lib.MRLIB_SetOfflineActivationKey.argtypes = [c_int64]
    lib.MRLIB_SetOfflineActivationKeyEx2.restype = c_bool
    lib.MRLIB_SetOfflineActivationKeyEx2.argtypes = [
        c_int64, c_int64, POINTER(c_int), POINTER(c_int), POINTER(c_int),
    ]


def _try_open(lib: ctypes.CDLL, path: str) -> tuple[bool, str]:
    model = (c_longlong * 1)()
    try:
        ok = bool(lib.MRLIB_OpenSimulationModel(path, model))
    except Exception as exc:
        return False, f"raised: {exc}"
    if not ok:
        return False, "MRLIB_OpenSimulationModel returned FALSE"
    var_count_or_neg = lib.MRLIB_GetModelDataLength(model[0], 0)
    note = f"model_ptr={model[0]!r}, GetModelDataLength(sim=0)={var_count_or_neg}"
    try:
        lib.MRLIB_CloseSimulationModel(model[0])
    except Exception:
        pass
    return True, note


def _find_recent_vmrs() -> str | None:
    """Look for a recent .vmrs file in common ModelRisk save locations."""
    bases = [
        Path.home() / "Documents",
        Path.home() / "Desktop",
        Path("C:/Users/timou/source/repos/ModelRisk/ModelRisk_Project/ModelRisk"),
    ]
    candidates: list[Path] = []
    for base in bases:
        if not base.is_dir():
            continue
        candidates.extend(base.glob("**/*.vmrs"))
    if not candidates:
        return None
    candidates.sort(key=lambda p: p.stat().st_mtime, reverse=True)
    return str(candidates[0])


def main(argv: list[str]) -> int:
    workbook_arg = argv[1] if len(argv) > 1 else None
    lines: list[str] = [
        "# MRService.dll spike",
        "",
        f"Ran at {datetime.now(UTC).isoformat()}.",
        f"Python bitness: {'64' if sys.maxsize > 2**32 else '32'}-bit.",
        "",
    ]

    dll_path = _find_dll()
    if dll_path is None:
        lines.append("**FAIL** — MRService.dll not found in any standard location:")
        for c in DLL_CANDIDATES:
            lines.append(f"  - `{c}`")
        OUT_PATH.write_text("\n".join(lines), encoding="utf-8")
        print("\n".join(lines))
        return 2

    lines.append(f"**DLL located** at `{dll_path}`.")
    lines.append("")
    try:
        lib = _load(dll_path)
        _setup_signatures(lib)
    except Exception as exc:
        lines.append(f"**FAIL** — LoadLibrary or signature setup: {exc}")
        OUT_PATH.write_text("\n".join(lines), encoding="utf-8")
        print("\n".join(lines))
        return 2
    lines.append("DLL loaded and signatures configured.")
    lines.append("")

    # Activation. Defaults to the single-key variant per the canonical
    # MRService.dll workflow; env-overridable so we don't bake keys into
    # the repo. Two-key variants are also available if the install needs
    # them.
    import os as _os
    activated = False
    key_single = _os.environ.get("MRSERVICE_ACTIVATION_KEY")
    if key_single:
        try:
            activated = bool(
                lib.MRLIB_SetOfflineActivationKey(c_int64(int(key_single)))
            )
            lines.append(
                f"Single-key activation: {'OK' if activated else 'returned False'}"
                f" (MRLIB_SetOfflineActivationKey)."
            )
        except Exception as exc:
            lines.append(f"Single-key activation raised: {exc}")
    if not activated:
        key1_env = _os.environ.get("MRSERVICE_ACTIVATION_KEY1")
        key2_env = _os.environ.get("MRSERVICE_ACTIVATION_KEY2")
        if key1_env and key2_env:
            year, month, day = c_int(), c_int(), c_int()
            try:
                activated = bool(
                    lib.MRLIB_SetOfflineActivationKeyEx2(
                        c_int64(int(key1_env)), c_int64(int(key2_env)),
                        byref(year), byref(month), byref(day),
                    )
                )
                if activated:
                    lines.append(
                        f"Two-key activation succeeded (license expires "
                        f"{year.value}-{month.value:02d}-{day.value:02d})."
                    )
                else:
                    lines.append("Two-key activation returned False.")
            except Exception as exc:
                lines.append(f"Two-key activation raised: {exc}")
    if not activated:
        lines.append(
            "**No valid activation key supplied.** Set "
            "MRSERVICE_ACTIVATION_KEY (single int64) or "
            "MRSERVICE_ACTIVATION_KEY1/2 (two int64s) before running."
        )
    lines.append("")

    # Try opening a workbook (the headless-mode question).
    lines.append("## Phase B test: open .xlsx headlessly")
    if workbook_arg:
        ok, note = _try_open(lib, workbook_arg)
        verdict = "VIABLE" if ok else "NOT VIABLE"
        lines.append(f"**{verdict}** — `{workbook_arg}` → {note}")
    else:
        lines.append(
            "_Skipped (no workbook path given on command line). "
            "Re-run as `python spike_mrservice.py path/to/workbook.xlsx`._"
        )
    lines.append("")

    # Try opening a .vmrs (the read-only path).
    lines.append("## Read-path test: open .vmrs results")
    vmrs = _find_recent_vmrs()
    if vmrs:
        ok, note = _try_open(lib, vmrs)
        verdict = "VIABLE" if ok else "NOT VIABLE"
        lines.append(f"**{verdict}** — `{vmrs}` → {note}")
    else:
        lines.append("_Skipped — no .vmrs file found in common locations._")
    lines.append("")

    lines.append("## Disposition")
    lines.append("")
    lines.append(
        "- If the .xlsx test is VIABLE, the v0.3 architecture is fully "
        "headless: MRService.dll opens the workbook, runs the simulation, "
        "and reads results — no Excel needed.\n"
        "- If only the .vmrs test is VIABLE, the v0.3 architecture is "
        "Excel-triggered + headless-read: the XLL in Excel runs the "
        "simulation, MRService.dll reads the resulting .vmrs.\n"
        "- If neither is VIABLE, the issue is likely activation "
        "(MRLIB_SetOfflineActivationKeyEx2 required) or a DLL "
        "dependency that didn't resolve."
    )

    OUT_PATH.write_text("\n".join(lines), encoding="utf-8")
    # stdout may be cp1251 on Windows; write file is the authoritative output.
    try:
        print("\n".join(lines))
    except UnicodeEncodeError:
        print(f"(report contains non-ASCII; see {OUT_PATH})")
    print(f"\nReport written to {OUT_PATH}")
    return 0


if __name__ == "__main__":
    raise SystemExit(main(sys.argv))

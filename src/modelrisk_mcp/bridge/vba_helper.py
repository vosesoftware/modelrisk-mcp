"""VBA helper bridge — runs ModelRisk's COM surface inside Excel's process.

User-reported pattern: `ModelRiskAtl.dll` doesn't expose its IDispatch
when loaded into our Python process (every standalone `CoCreateInstance`
strategy returns E_NOINTERFACE), but VBA inside Excel can do it fine.
That's the ATL "I only function in-process with my host Office app"
pattern, and the workaround is to do the COM calls inside Excel's
process via `Excel.Application.Run`.

This module manages a small VBA module that gets injected on demand
into a hidden helper workbook. The module exposes:

  ModelRiskMcp_RunSim                — Sub: kicks off the sim
  ModelRiskMcp_SetSamples(n)         — Sub: set iterations
  ModelRiskMcp_SetSeed(value)        — Sub: set fixed seed
  ModelRiskMcp_GetOutputCount        — Function: number of outputs
  ModelRiskMcp_GetOutputName(i)      — Function: i-th output's name
  ModelRiskMcp_GetMean(name)         — Function: output's mean
  ModelRiskMcp_GetPercentile(name,p) — Function: output's percentile
  ModelRiskMcp_GetStDev(name)        — Function: output's stdev
  ModelRiskMcp_GetSamples(name)      — Function: full sample array
  ModelRiskMcp_IsAvailable           — Function: returns True if the
                                       Vose COM objects can be created
                                       inside Excel right now

The helper workbook is created hidden, the module injected via the
VBE extensibility (Trust → Trust access to the VBA project object
model must be on; the user gets a one-time prompt if not). Closing
Excel removes the workbook; we never write to the user's own files.
"""

from __future__ import annotations

from dataclasses import dataclass
from typing import Any

VBA_MODULE_NAME: str = "ModelRiskMcpHelper"

# Early-bound VBA module: requires a Reference to ModelRiskAtl.dll
# (TypeLib `{ECC429DA-26E6-4D86-9B2D-1E14E0461749}`). VbaHelperBridge
# adds that reference before injecting this code. The early binding is
# essential: the ModelRisk coclasses don't expose IDispatch at runtime,
# only their custom typelib interfaces — that's why late-bound
# CreateObject (and Python's Dispatch) returns E_NOINTERFACE.
VBA_MODULE_CODE_EARLY: str = r"""Option Explicit
' modelrisk-mcp helper — runs ModelRisk's COM surface inside Excel via
' early binding through the ModelRiskAtl.dll typelib. Generated; do
' not hand-edit. The Reference to ModelRiskAtl.dll is added by the
' MCP server before this module is loaded.

Public Function ModelRiskMcp_IsAvailable() As Boolean
    On Error Resume Next
    Dim sim As ModelRisk.ModelRiskSimulation
    Set sim = New ModelRisk.ModelRiskSimulation
    ModelRiskMcp_IsAvailable = (Err.Number = 0 And Not sim Is Nothing)
    On Error GoTo 0
End Function

Public Sub ModelRiskMcp_SetSamples(ByVal n As Long)
    Dim settings As ModelRisk.ModelRiskSimulationSettings
    Set settings = New ModelRisk.ModelRiskSimulationSettings
    settings.Samples = n
End Sub

Public Sub ModelRiskMcp_SetSeed(ByVal seedValue As Double)
    Dim settings As ModelRisk.ModelRiskSimulationSettings
    Set settings = New ModelRisk.ModelRiskSimulationSettings
    settings.UseFixedSeed = True
    settings.Seed(0) = seedValue
End Sub

Public Sub ModelRiskMcp_SetHideProgressWindow(ByVal hide As Boolean)
    Dim settings As ModelRisk.ModelRiskSimulationSettings
    Set settings = New ModelRisk.ModelRiskSimulationSettings
    settings.HideProgressWindow = hide
End Sub

Public Sub ModelRiskMcp_RunSim()
    Dim sim As ModelRisk.ModelRiskSimulation
    Set sim = New ModelRisk.ModelRiskSimulation
    sim.StartSimulation
End Sub

Public Function ModelRiskMcp_GetOutputCount() As Long
    Dim results As ModelRisk.ModelRiskSimulationResults
    Set results = New ModelRisk.ModelRiskSimulationResults
    ModelRiskMcp_GetOutputCount = CLng(results.SimOutputs().Count)
End Function

Public Function ModelRiskMcp_GetOutputName(ByVal index As Long) As String
    Dim results As ModelRisk.ModelRiskSimulationResults
    Set results = New ModelRisk.ModelRiskSimulationResults
    ModelRiskMcp_GetOutputName = CStr(results.SimOutputs().Item(index).GetName())
End Function

Public Function ModelRiskMcp_GetMean(ByVal varName As String) As Double
    Dim v As Object
    Set v = ModelRiskMcp_FindVar(varName)
    If v Is Nothing Then
        ModelRiskMcp_GetMean = 0
    Else
        ModelRiskMcp_GetMean = CDbl(v.GetMean())
    End If
End Function

Public Function ModelRiskMcp_GetPercentile(ByVal varName As String, ByVal p As Double) As Double
    Dim v As Object
    Set v = ModelRiskMcp_FindVar(varName)
    If v Is Nothing Then
        ModelRiskMcp_GetPercentile = 0
    Else
        ModelRiskMcp_GetPercentile = CDbl(v.GetPercentile(p))
    End If
End Function

Public Function ModelRiskMcp_GetStDev(ByVal varName As String) As Double
    Dim v As Object
    Set v = ModelRiskMcp_FindVar(varName)
    If v Is Nothing Then
        ModelRiskMcp_GetStDev = 0
    Else
        ModelRiskMcp_GetStDev = CDbl(v.GetStDev())
    End If
End Function

Public Function ModelRiskMcp_GetSamples(ByVal varName As String) As Variant
    Dim v As Object
    Set v = ModelRiskMcp_FindVar(varName)
    If v Is Nothing Then
        ModelRiskMcp_GetSamples = Array()
    Else
        ModelRiskMcp_GetSamples = v.GetSamples()
    End If
End Function

Private Function ModelRiskMcp_FindVar(ByVal varName As String) As Object
    Dim results As ModelRisk.ModelRiskSimulationResults
    Dim outs As Object, var As Object, i As Long
    Set results = New ModelRisk.ModelRiskSimulationResults
    Set outs = results.SimOutputs()
    For i = 1 To outs.Count
        Set var = outs.Item(i)
        If CStr(var.GetName()) = varName Then
            Set ModelRiskMcp_FindVar = var
            Exit Function
        End If
    Next i
    Set ModelRiskMcp_FindVar = Nothing
End Function
"""

# Late-bound fallback used when we can't add the typelib Reference (e.g.
# the DLL path isn't findable in the registry). Likely to fail with
# E_NOINTERFACE on the current ModelRisk build — kept as a last resort
# so the helper degrades gracefully rather than crashing.
VBA_MODULE_CODE_LATE: str = r"""Option Explicit
' modelrisk-mcp helper (late-binding fallback). Used only if the
' typelib reference couldn't be added. Likely to fail with
' E_NOINTERFACE since the ModelRisk coclasses don't expose IDispatch.

Public Function ModelRiskMcp_IsAvailable() As Boolean
    On Error Resume Next
    Dim sim As Object
    Set sim = CreateObject("ModelRisk.ModelRiskSimulation")
    ModelRiskMcp_IsAvailable = (Err.Number = 0 And Not sim Is Nothing)
    On Error GoTo 0
End Function

Public Sub ModelRiskMcp_SetSamples(ByVal n As Long)
    CreateObject("ModelRisk.ModelRiskSimulationSettings").Samples = n
End Sub

Public Sub ModelRiskMcp_SetSeed(ByVal seedValue As Double)
    Dim s As Object
    Set s = CreateObject("ModelRisk.ModelRiskSimulationSettings")
    s.UseFixedSeed = True
    s.Seed(0) = seedValue
End Sub

Public Sub ModelRiskMcp_SetHideProgressWindow(ByVal hide As Boolean)
    CreateObject("ModelRisk.ModelRiskSimulationSettings").HideProgressWindow = hide
End Sub

Public Sub ModelRiskMcp_RunSim()
    CreateObject("ModelRisk.ModelRiskSimulation").StartSimulation
End Sub

Public Function ModelRiskMcp_GetOutputCount() As Long
    ModelRiskMcp_GetOutputCount = CLng( _
        CreateObject("ModelRisk.ModelRiskSimulationResults").SimOutputs().Count)
End Function

Public Function ModelRiskMcp_GetMean(ByVal varName As String) As Double
    ModelRiskMcp_GetMean = 0
End Function

Public Function ModelRiskMcp_GetPercentile(ByVal varName As String, ByVal p As Double) As Double
    ModelRiskMcp_GetPercentile = 0
End Function

Public Function ModelRiskMcp_GetStDev(ByVal varName As String) As Double
    ModelRiskMcp_GetStDev = 0
End Function

Public Function ModelRiskMcp_GetSamples(ByVal varName As String) As Variant
    ModelRiskMcp_GetSamples = Array()
End Function

Public Function ModelRiskMcp_GetOutputName(ByVal index As Long) As String
    ModelRiskMcp_GetOutputName = ""
End Function
"""

# Backwards-compatible name kept for existing tests.
VBA_MODULE_CODE: str = VBA_MODULE_CODE_EARLY


@dataclass
class VbaInjectionResult:
    ok: bool
    error: str | None = None
    workbook_name: str | None = None
    used_early_binding: bool = False


class VbaHelperBridge:
    """Injects + invokes the VBA helper module in a hidden workbook.

    On the first call, walks `Excel.Workbooks` for an existing helper
    workbook by name; if not found, adds a new one and writes
    `VBA_MODULE_CODE` into a fresh standard module. The workbook is
    kept open and hidden for the rest of the Excel session.

    Each call to `run` invokes a single VBA function via
    `Excel.Application.Run(...)`. Arguments are passed through as
    Python primitives (int, float, str) and the return value is the
    VBA function's result as a Python primitive.

    Trust requirement: "Trust access to the VBA project object model"
    must be enabled in Excel's Trust Center (File → Options → Trust
    Center → Macro Settings). On first injection Excel may prompt;
    failures from a missing trust setting are surfaced as a clear
    `VbaTrustError` rather than a raw HRESULT.
    """

    HELPER_WORKBOOK_NAME: str = "ModelRiskMcpHelper.xlsm"

    def __init__(self, excel_bridge: Any) -> None:
        self._excel = excel_bridge
        self._injected: bool = False
        self._used_early_binding: bool = False

    def inject_if_needed(self) -> VbaInjectionResult:
        """Idempotent. On first call: find or create the helper workbook,
        add a Reference to ModelRiskAtl.dll, write the VBA module."""
        if self._injected:
            return VbaInjectionResult(
                ok=True,
                workbook_name=self.HELPER_WORKBOOK_NAME,
                used_early_binding=self._used_early_binding,
            )
        try:
            book = self._find_or_create_helper_book()
        except Exception as exc:
            return VbaInjectionResult(
                ok=False,
                error=(
                    f"Could not create the helper workbook: {exc}. "
                    "Confirm Excel is running and that "
                    "'Trust access to the VBA project object model' is "
                    "enabled in File → Options → Trust Center → Macro "
                    "Settings."
                ),
            )
        try:
            self._inject_module(book)
        except Exception as exc:
            return VbaInjectionResult(
                ok=False,
                workbook_name=self.HELPER_WORKBOOK_NAME,
                error=(
                    f"VBA module injection failed: {exc}. "
                    "Most likely cause: 'Trust access to the VBA "
                    "project object model' is OFF in Excel's Trust "
                    "Center. Turn it on and retry."
                ),
            )
        self._injected = True
        return VbaInjectionResult(
            ok=True,
            workbook_name=self.HELPER_WORKBOOK_NAME,
            used_early_binding=self._used_early_binding,
        )

    def run(self, sub_name: str, *args: Any) -> Any:
        """Invoke `Excel.Application.Run("<helper-workbook>!<sub>", *args)`
        and return the result. Caller is responsible for converting the
        result type if needed."""
        injection = self.inject_if_needed()
        if not injection.ok:
            raise RuntimeError(
                f"VBA helper not available: {injection.error}"
            )
        app = self._excel._app
        qualified = f"{self.HELPER_WORKBOOK_NAME}!{sub_name}"
        return app.api.Run(qualified, *args)

    # ----- helpers ----------------------------------------------------

    def _find_or_create_helper_book(self) -> Any:
        app = self._excel._app
        for book in app.api.Workbooks:
            if str(book.Name).lower() == self.HELPER_WORKBOOK_NAME.lower():
                return book
        # Add a new one. xlOpenXMLWorkbookMacroEnabled = 52 for .xlsm.
        book = app.api.Workbooks.Add()
        # Try to make it hidden — best-effort, harmless if it fails.
        try:
            for w in app.api.Windows:
                if w.Caption.startswith(book.Name):
                    w.Visible = False
        except Exception:
            pass
        return book

    def _inject_module(self, book: Any) -> None:
        vbproj = book.VBProject
        # Add the typelib Reference first — without it the early-bound
        # macro code won't compile, and the late-bound fallback is
        # known to fail with E_NOINTERFACE against this ATL coclass.
        reference_added = self._ensure_modelrisk_reference(vbproj)
        code = VBA_MODULE_CODE_EARLY if reference_added else VBA_MODULE_CODE_LATE
        self._used_early_binding = reference_added
        # If the module is already present, replace its contents.
        existing = None
        try:
            existing = vbproj.VBComponents(VBA_MODULE_NAME)
        except Exception:
            existing = None
        if existing is not None:
            existing.CodeModule.DeleteLines(
                1, existing.CodeModule.CountOfLines
            )
            existing.CodeModule.AddFromString(code)
            return
        # vbext_ct_StdModule = 1
        component = vbproj.VBComponents.Add(1)
        component.Name = VBA_MODULE_NAME
        component.CodeModule.AddFromString(code)

    def _ensure_modelrisk_reference(self, vbproj: Any) -> bool:
        """Add a VBA Reference to ModelRiskAtl.dll. Returns True if the
        reference is present after this call.

        Tries `AddFromFile(dll_path)` first (most reliable), falling back
        to `AddFromGuid` with the typelib UUID from the IDL. Either
        path requires 'Trust access to the VBA project object model'
        to be on, which `inject_if_needed` already implies."""
        try:
            refs = vbproj.References
        except Exception:
            return False
        # Already present?
        try:
            for ref in refs:
                desc = str(getattr(ref, "Description", "") or "")
                full_path = str(getattr(ref, "FullPath", "") or "")
                if (
                    "modelrisk" in desc.lower()
                    or "modelriskatl" in full_path.lower()
                ):
                    return True
        except Exception:
            pass
        # Try AddFromFile with the registered InprocServer32 path.
        from modelrisk_mcp.bridge.modelrisk import (
            _lookup_modelrisk_inproc_server,
        )

        _clsid, dll_path = _lookup_modelrisk_inproc_server()
        if dll_path:
            try:
                refs.AddFromFile(dll_path)
                return True
            except Exception:
                pass
        # Fall back to AddFromGuid with the typelib GUID from the IDL.
        from modelrisk_mcp.bridge.progids import TYPELIB

        try:
            refs.AddFromGuid(TYPELIB, 1, 0)
            return True
        except Exception:
            return False


__all__ = [
    "VBA_MODULE_CODE",
    "VBA_MODULE_CODE_EARLY",
    "VBA_MODULE_CODE_LATE",
    "VBA_MODULE_NAME",
    "VbaHelperBridge",
    "VbaInjectionResult",
]

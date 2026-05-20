"""Tests for the VBA helper bridge.

Excel COM is mocked; we verify:
- VBA module text contains every Mcp_* entry point the bridge calls
- inject_if_needed creates the helper workbook if absent, finds it if present
- run() formats the qualified name correctly and forwards args
- inject_if_needed is idempotent
- a clear error surfaces when VBOM trust is off
"""

from __future__ import annotations

from typing import Any

import pytest

from modelrisk_mcp.bridge.vba_helper import (
    VBA_MODULE_CODE,
    VBA_MODULE_NAME,
    VbaHelperBridge,
)


class _FakeVBComponents:
    def __init__(self) -> None:
        self._items: list[_FakeVBComponent] = []

    def Add(self, kind: int) -> _FakeVBComponent:  # noqa: N802
        comp = _FakeVBComponent(name="NewComponent", kind=kind)
        self._items.append(comp)
        return comp

    def __call__(self, name: str) -> _FakeVBComponent:
        # Look up by the component's *current* name. The bridge code
        # renames after Add(), so lookups happen against the post-
        # rename name.
        for comp in self._items:
            if comp.name == name:
                return comp
        raise KeyError(name)


class _FakeVBComponent:
    def __init__(self, name: str, kind: int = 1) -> None:
        self.name = name
        self.kind = kind
        self.CodeModule = _FakeCodeModule()

    @property
    def Name(self) -> str:  # noqa: N802
        return self.name

    @Name.setter
    def Name(self, value: str) -> None:  # noqa: N802
        self.name = value


class _FakeCodeModule:
    def __init__(self) -> None:
        self.code: str = ""

    @property
    def CountOfLines(self) -> int:  # noqa: N802
        return len(self.code.splitlines())

    def AddFromString(self, text: str) -> None:  # noqa: N802
        self.code += text

    def DeleteLines(self, start: int, count: int) -> None:  # noqa: N802
        self.code = ""


class _FakeReferences:
    def __init__(self, *, allow_add: bool = True) -> None:
        self._refs: list[dict[str, str]] = []
        self._allow_add = allow_add

    def __iter__(self):
        for r in self._refs:
            yield type("_Ref", (), r)()

    def AddFromFile(self, path: str) -> None:  # noqa: N802
        if not self._allow_add:
            raise RuntimeError("AddFromFile blocked")
        self._refs.append({"Description": "ModelRisk", "FullPath": path})

    def AddFromGuid(self, guid: str, major: int, minor: int) -> None:  # noqa: N802
        if not self._allow_add:
            raise RuntimeError("AddFromGuid blocked")
        self._refs.append(
            {"Description": f"ModelRisk ({guid})", "FullPath": ""}
        )


class _FakeWorkbook:
    def __init__(self, name: str, *, allow_reference_add: bool = True) -> None:
        self.Name = name
        # VBProject is a tiny shim with both VBComponents and References.
        self.VBProject = type(
            "_VBProj",
            (),
            {
                "VBComponents": _FakeVBComponents(),
                "References": _FakeReferences(allow_add=allow_reference_add),
            },
        )()


class _FakeWorkbooks:
    def __init__(self) -> None:
        self._books: list[_FakeWorkbook] = []

    def __iter__(self):
        return iter(self._books)

    def Add(self) -> _FakeWorkbook:  # noqa: N802
        book = _FakeWorkbook(name="ModelRiskMcpHelper.xlsm")
        self._books.append(book)
        return book


class _FakeApp:
    def __init__(self) -> None:
        self.Workbooks = _FakeWorkbooks()
        self.Windows: list[Any] = []
        self.run_calls: list[tuple[str, tuple[Any, ...]]] = []

    def Run(self, qualified: str, *args: Any) -> Any:  # noqa: N802
        self.run_calls.append((qualified, args))
        return "ok"


class _FakeExcelBridge:
    def __init__(self) -> None:
        self._app = type("_AppShim", (), {})()
        self._app.api = _FakeApp()

    def connect(self) -> None:
        pass


# ----------------------------------------------------------------------


class TestVbaModuleCode:
    def test_module_defines_every_entry_point(self) -> None:
        for name in (
            "ModelRiskMcp_IsAvailable",
            "ModelRiskMcp_SetSamples",
            "ModelRiskMcp_SetSeed",
            "ModelRiskMcp_SetHideProgressWindow",
            "ModelRiskMcp_RunSim",
            "ModelRiskMcp_GetOutputCount",
            "ModelRiskMcp_GetOutputName",
            "ModelRiskMcp_GetMean",
            "ModelRiskMcp_GetPercentile",
            "ModelRiskMcp_GetStDev",
            "ModelRiskMcp_GetSamples",
        ):
            assert name in VBA_MODULE_CODE, f"{name} missing from VBA module"

    def test_module_name_is_consistent(self) -> None:
        assert VBA_MODULE_NAME == "ModelRiskMcpHelper"


class TestInjection:
    def test_creates_workbook_when_absent(self) -> None:
        excel = _FakeExcelBridge()
        helper = VbaHelperBridge(excel)
        result = helper.inject_if_needed()
        assert result.ok is True
        assert result.workbook_name == "ModelRiskMcpHelper.xlsm"
        books = list(excel._app.api.Workbooks._books)
        assert len(books) == 1
        # Module body landed in the book's project.
        comp = books[0].VBProject.VBComponents("ModelRiskMcpHelper")
        assert "ModelRiskMcp_RunSim" in comp.CodeModule.code

    def test_idempotent_when_already_injected(self) -> None:
        excel = _FakeExcelBridge()
        helper = VbaHelperBridge(excel)
        helper.inject_if_needed()
        helper.inject_if_needed()
        # Still just one workbook.
        assert len(excel._app.api.Workbooks._books) == 1

    def test_reuses_existing_helper_workbook(self) -> None:
        excel = _FakeExcelBridge()
        # Pre-add a workbook with the helper name (mimics a previous
        # session's leftover).
        pre = _FakeWorkbook("ModelRiskMcpHelper.xlsm")
        excel._app.api.Workbooks._books.append(pre)
        helper = VbaHelperBridge(excel)
        helper.inject_if_needed()
        assert len(excel._app.api.Workbooks._books) == 1
        # Module code went into the existing book.
        assert (
            "ModelRiskMcp_RunSim"
            in pre.VBProject.VBComponents("ModelRiskMcpHelper").CodeModule.code
        )

    def test_uses_early_binding_when_reference_added(
        self, monkeypatch: pytest.MonkeyPatch
    ) -> None:
        """When VBProject.References.AddFromFile succeeds, the early-
        bound macro code is injected (uses `As ModelRisk.<Class>`
        types). This is the path that should actually work against a
        custom-interface-only coclass."""
        # Make the registry lookup return a path we can pretend is the DLL.
        monkeypatch.setattr(
            "modelrisk_mcp.bridge.modelrisk._lookup_modelrisk_inproc_server",
            lambda: ("{570013C9-...}", "C:/fake/ModelRiskAtl.dll"),
        )
        excel = _FakeExcelBridge()
        helper = VbaHelperBridge(excel)
        result = helper.inject_if_needed()
        assert result.ok is True
        assert result.used_early_binding is True
        books = list(excel._app.api.Workbooks._books)
        comp = books[0].VBProject.VBComponents("ModelRiskMcpHelper")
        # Early-bound code uses `Dim sim As ModelRisk.ModelRiskSimulation`
        # — that string isn't in the late-bound fallback.
        assert "As ModelRisk.ModelRiskSimulation" in comp.CodeModule.code

    def test_falls_back_to_late_binding_when_reference_add_fails(
        self, monkeypatch: pytest.MonkeyPatch
    ) -> None:
        """If AddFromFile and AddFromGuid both fail, the helper still
        injects the late-bound fallback so subsequent calls don't
        crash. The fallback is documented to likely fail with
        E_NOINTERFACE — that's a downstream concern."""
        monkeypatch.setattr(
            "modelrisk_mcp.bridge.modelrisk._lookup_modelrisk_inproc_server",
            lambda: (None, None),
        )
        # And block AddFromGuid too.
        excel = _FakeExcelBridge()
        book = _FakeWorkbook(
            "ModelRiskMcpHelper.xlsm", allow_reference_add=False
        )
        excel._app.api.Workbooks._books.append(book)
        helper = VbaHelperBridge(excel)
        result = helper.inject_if_needed()
        assert result.ok is True
        assert result.used_early_binding is False
        comp = book.VBProject.VBComponents("ModelRiskMcpHelper")
        # Late-bound code uses CreateObject; early-bound does not.
        assert "CreateObject" in comp.CodeModule.code
        assert "As ModelRisk.ModelRiskSimulation" not in comp.CodeModule.code

    def test_vbom_trust_failure_surfaces_clear_error(self) -> None:
        """Excel raises a COM error when 'Trust access to the VBA
        project object model' is off; we should surface the actionable
        message."""
        excel = _FakeExcelBridge()

        # Override the workbook to raise on VBProject access.
        class _BlockedVBProj:
            @property
            def VBComponents(self):  # noqa: N802
                raise RuntimeError("VBOM trust off")

        broken = _FakeWorkbook("ModelRiskMcpHelper.xlsm")
        broken.VBProject = _BlockedVBProj()
        excel._app.api.Workbooks._books.append(broken)

        helper = VbaHelperBridge(excel)
        result = helper.inject_if_needed()
        assert result.ok is False
        assert "Trust access to the VBA project object model" in (
            result.error or ""
        )


class TestRun:
    def test_qualifies_macro_name_with_workbook(self) -> None:
        excel = _FakeExcelBridge()
        helper = VbaHelperBridge(excel)
        helper.run("ModelRiskMcp_SetSamples", 1000)
        assert excel._app.api.run_calls == [
            ("ModelRiskMcpHelper.xlsm!ModelRiskMcp_SetSamples", (1000,))
        ]

    def test_forwards_positional_args(self) -> None:
        excel = _FakeExcelBridge()
        helper = VbaHelperBridge(excel)
        helper.run("ModelRiskMcp_GetPercentile", "Profit", 0.9)
        assert excel._app.api.run_calls[-1] == (
            "ModelRiskMcpHelper.xlsm!ModelRiskMcp_GetPercentile",
            ("Profit", 0.9),
        )

    def test_raises_when_injection_fails(self) -> None:
        excel = _FakeExcelBridge()

        # Force inject_if_needed to fail by making VBComponents raise.
        class _BlockedVBProj:
            @property
            def VBComponents(self):  # noqa: N802
                raise RuntimeError("blocked")

        excel._app.api.Workbooks._books.append(
            type("BlockedBook", (), {
                "Name": "ModelRiskMcpHelper.xlsm",
                "VBProject": _BlockedVBProj(),
            })()
        )
        helper = VbaHelperBridge(excel)
        with pytest.raises(RuntimeError, match="VBA helper not available"):
            helper.run("ModelRiskMcp_SetSamples", 1)

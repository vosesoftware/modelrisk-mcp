"""Phase 3 building-tool tests with a mocked bridge.

Acceptance criteria from spec §13 Phase 3:
- dry_run=True (the default) does not mutate any workbook.
- dry_run=False writes through ExcelBridge.write_cell.
- bulk-write guard at >50 cells; time-series and copula are exempt.
- non-Vose formula refusal works via formula-tokenised detection.
- writes.log gains one record per committed write.
- restore_cell round-trip recovers the pre-write state.
"""

from __future__ import annotations

import json
from collections.abc import Iterator
from pathlib import Path
from typing import Any

import pytest

from modelrisk_mcp.bridge.catalogue import load_catalogue
from modelrisk_mcp.bridge.modelrisk import ModelRiskBridge
from modelrisk_mcp.config import Settings
from modelrisk_mcp.errors import CellReferenceError, UnknownFunctionError
from modelrisk_mcp.safety import WriterMutex
from modelrisk_mcp.schemas.workbook import CellInfo, CellRef, WorkbookInfo
from modelrisk_mcp.tools import building, reading, restore

# ----------------------------------------------------------------------
# Fakes
# ----------------------------------------------------------------------


class WritableFakeExcel:
    """In-memory Excel substitute that records cell writes and can be
    introspected by tests."""

    def __init__(self) -> None:
        self.cells: dict[tuple[str, str, str], CellInfo] = {}
        self.named_ranges: dict[tuple[str, str], str] = {}
        self.undo_calls = 0

    # --- ExcelBridge surface ---------------------------------------

    def list_workbooks(self) -> list[WorkbookInfo]:
        names = {wb for (wb, _, _) in self.cells} or {"book.xlsx"}
        return [
            WorkbookInfo(name=n, path=f"C:/{n}", sheets=["Sheet1"])
            for n in names
        ]

    def get_active_workbook(self) -> WorkbookInfo:
        return self.list_workbooks()[0]

    def get_cell(self, workbook: str, sheet: str, cell: str) -> CellInfo:
        key = (workbook, sheet, cell.upper())
        if key in self.cells:
            return self.cells[key]
        return CellInfo(
            ref=CellRef(workbook=workbook, sheet=sheet, cell=cell.upper()),
            formula="",
            value=None,
            cell_type="empty",
        )

    def write_cell(
        self, workbook: str, sheet: str, cell: str, formula: str
    ) -> None:
        ref = CellRef(workbook=workbook, sheet=sheet, cell=cell.upper())
        self.cells[(workbook, sheet, cell.upper())] = CellInfo(
            ref=ref,
            formula=formula,
            value=None,
            cell_type="formula" if formula else "empty",
        )

    def write_range(
        self,
        workbook: str,
        sheet: str,
        range_ref: str,
        formulas: list[list[str]],
    ) -> None:
        # Not used by Phase 3 building tools (they all single-cell write).
        pass

    def read_range(self, *args: Any, **kwargs: Any) -> Any:
        from modelrisk_mcp.schemas.workbook import RangeInfo
        return RangeInfo(
            workbook=args[0] if args else "", sheet="", range_ref="A1"
        )

    def iterate_cells(
        self,
        workbook: str,
        predicate: Any = None,
        *,
        sheet: str | None = None,
    ) -> Iterator[CellInfo]:
        for (wb, sh, _cell), cellinfo in self.cells.items():
            if wb != workbook:
                continue
            if sheet is not None and sh != sheet:
                continue
            if predicate is None or predicate(cellinfo):
                yield cellinfo

    def set_named_range(
        self, workbook: str, name: str, range_ref: str
    ) -> None:
        self.named_ranges[(workbook, name)] = range_ref

    def undo(self) -> None:
        self.undo_calls += 1


# ----------------------------------------------------------------------
# Fixture
# ----------------------------------------------------------------------


@pytest.fixture
def fake_excel() -> WritableFakeExcel:
    return WritableFakeExcel()


@pytest.fixture
def bridge_with_audit(
    fake_excel: WritableFakeExcel, tmp_path: Path
) -> Iterator[tuple[ModelRiskBridge, Path]]:
    audit_log = tmp_path / "writes.log"
    settings = Settings(log_dir=tmp_path, writes_log_name="writes.log")
    bridge = ModelRiskBridge(
        excel=fake_excel,  # type: ignore[arg-type]
        catalogue=load_catalogue(),
        settings=settings,
        writer_mutex=WriterMutex(name="modelrisk-mcp-test-phase3"),
    )
    reading.set_bridge_for_testing(bridge)
    yield bridge, audit_log
    reading.set_bridge_for_testing(None)


# ----------------------------------------------------------------------
# dry_run = True (the default) — must not mutate
# ----------------------------------------------------------------------


class TestDryRunDoesNotMutate:
    """Spec §13 Phase 3 acceptance: default dry_run=True is verified not
    to mutate any workbook for every building tool."""

    def test_insert_distribution_dry_run(
        self,
        fake_excel: WritableFakeExcel,
        bridge_with_audit: tuple[ModelRiskBridge, Path],
    ) -> None:
        result = building.insert_distribution(
            "book.xlsx", "Sheet1", "B1",
            "VoseNormal",
            [{"name": "mu", "value": 0}, {"name": "sigma", "value": 1}],
        )
        assert result.written is False
        assert result.formula == "=VoseNormal(0,1)"
        assert ("book.xlsx", "Sheet1", "B1") not in fake_excel.cells

    def test_wrap_with_input_dry_run(
        self,
        fake_excel: WritableFakeExcel,
        bridge_with_audit: tuple[ModelRiskBridge, Path],
    ) -> None:
        fake_excel.write_cell("book.xlsx", "Sheet1", "B1", "=VoseNormal(0,1)")
        result = building.wrap_with_input(
            "book.xlsx", "Sheet1", "B1", name="Demand"
        )
        assert result.written is False
        assert result.formula == '=VoseInput("Demand")+VoseNormal(0,1)'
        # Original formula preserved.
        assert (
            fake_excel.cells[("book.xlsx", "Sheet1", "B1")].formula
            == "=VoseNormal(0,1)"
        )

    def test_wrap_with_output_dry_run(
        self,
        fake_excel: WritableFakeExcel,
        bridge_with_audit: tuple[ModelRiskBridge, Path],
    ) -> None:
        fake_excel.write_cell("book.xlsx", "Sheet1", "B1", "=A1+1")
        result = building.wrap_with_output(
            "book.xlsx", "Sheet1", "B1", name="Profit"
        )
        assert result.written is False
        assert result.formula.startswith('=VoseOutput("Profit")+')

    def test_replace_constant_dry_run(
        self,
        fake_excel: WritableFakeExcel,
        bridge_with_audit: tuple[ModelRiskBridge, Path],
    ) -> None:
        result = building.replace_constant_with_distribution(
            "book.xlsx", "Sheet1", "B1",
            "VoseNormal",
            [{"name": "mu", "value": 100}, {"name": "sigma", "value": 10}],
            input_name="Demand",
        )
        assert result.written is False
        assert '=VoseInput("Demand")+VoseNormal(100,10)' == result.formula
        assert ("book.xlsx", "Sheet1", "B1") not in fake_excel.cells

    def test_fit_distribution_dry_run(
        self,
        fake_excel: WritableFakeExcel,
        bridge_with_audit: tuple[ModelRiskBridge, Path],
    ) -> None:
        result = building.fit_distribution_to_data(
            "book.xlsx", "Sheet1", "B1",
            data_range="A1:A100",
            family="Beta",
        )
        assert result.written is False
        assert result.formula.startswith("=VoseBetaFit(")

    def test_create_aggregate_mc_dry_run(
        self,
        fake_excel: WritableFakeExcel,
        bridge_with_audit: tuple[ModelRiskBridge, Path],
    ) -> None:
        result = building.create_aggregate_mc(
            "book.xlsx", "Sheet1", "C1",
            frequency_object_cell="A1",
            severity_object_cell="B1",
        )
        assert result.written is False
        assert result.formula == "=VoseAggregateMC(A1,B1)"

    def test_create_risk_event_dry_run(
        self,
        fake_excel: WritableFakeExcel,
        bridge_with_audit: tuple[ModelRiskBridge, Path],
    ) -> None:
        # Use an object function from the catalogue.
        cat = bridge_with_audit[0].catalogue
        # Pick the first 'object' function whose required params we can satisfy.
        for spec in cat.filter("object"):
            required = [p for p in spec.parameters if p.required]
            if len(required) <= 2:
                impact_function = spec.name
                impact_params = [
                    {"name": p.name, "value": 1.0} for p in required
                ]
                break
        else:
            pytest.skip("No simple object function found in catalogue.")
        result = building.create_risk_event(
            "book.xlsx", "Sheet1", "D1",
            probability=0.1,
            impact_function_name=impact_function,
            impact_parameters=impact_params,
        )
        assert result.written is False
        assert result.formula.startswith("=VoseRiskEvent(0.1,")

    def test_create_time_series_dry_run(
        self,
        fake_excel: WritableFakeExcel,
        bridge_with_audit: tuple[ModelRiskBridge, Path],
    ) -> None:
        result = building.create_time_series(
            "book.xlsx", "Sheet1", "D2:D101",
            function_name="VoseTimeGBM",
            parameters=[
                {"name": "OutputSize", "value": 100},
                {"name": "mu", "value": 0.05},
                {"name": "sigma", "value": 0.2},
            ],
        )
        assert result.written is False
        assert result.formula.startswith("=VoseTimeGBM(")

    def test_create_copula_dry_run(
        self,
        fake_excel: WritableFakeExcel,
        bridge_with_audit: tuple[ModelRiskBridge, Path],
    ) -> None:
        result = building.create_copula(
            "book.xlsx", "Sheet1", "E1:E2",
            function_name="VoseCopulaMultiNormal",
            parameters=[{"name": "cov_matrix", "value": "F1:G2"}],
        )
        assert result.written is False
        assert result.formula == "=VoseCopulaMultiNormal(F1:G2)"

    def test_set_named_range_dry_run(
        self,
        fake_excel: WritableFakeExcel,
        bridge_with_audit: tuple[ModelRiskBridge, Path],
    ) -> None:
        result = building.set_named_range(
            "book.xlsx", "Demand", "Sheet1!$A$1:$A$10"
        )
        assert result["written"] == "false"
        assert ("book.xlsx", "Demand") not in fake_excel.named_ranges


# ----------------------------------------------------------------------
# dry_run = False — must mutate and log
# ----------------------------------------------------------------------


class TestExplicitWrite:
    def test_insert_distribution_writes(
        self,
        fake_excel: WritableFakeExcel,
        bridge_with_audit: tuple[ModelRiskBridge, Path],
    ) -> None:
        _, audit_log = bridge_with_audit
        result = building.insert_distribution(
            "book.xlsx", "Sheet1", "B1",
            "VoseNormal",
            [{"name": "mu", "value": 0}, {"name": "sigma", "value": 1}],
            dry_run=False,
        )
        assert result.written is True
        assert (
            fake_excel.cells[("book.xlsx", "Sheet1", "B1")].formula
            == "=VoseNormal(0,1)"
        )
        # Audit log written.
        records = [
            json.loads(line)
            for line in audit_log.read_text(encoding="utf-8").splitlines()
            if line.strip()
        ]
        assert len(records) == 1
        assert records[0]["after_formula"] == "=VoseNormal(0,1)"
        assert records[0]["cell"] == "book.xlsx!Sheet1!B1"

    def test_set_named_range_writes(
        self,
        fake_excel: WritableFakeExcel,
        bridge_with_audit: tuple[ModelRiskBridge, Path],
    ) -> None:
        result = building.set_named_range(
            "book.xlsx", "Demand", "Sheet1!$A$1:$A$10", dry_run=False
        )
        assert result["written"] == "true"
        assert (
            fake_excel.named_ranges[("book.xlsx", "Demand")]
            == "Sheet1!$A$1:$A$10"
        )


# ----------------------------------------------------------------------
# Safety: non-Vose-formula refusal (§11.5)
# ----------------------------------------------------------------------


class TestNonVoseRefusal:
    def test_refuses_to_overwrite_sum(
        self,
        fake_excel: WritableFakeExcel,
        bridge_with_audit: tuple[ModelRiskBridge, Path],
    ) -> None:
        fake_excel.write_cell("book.xlsx", "Sheet1", "B1", "=SUM(A1:A10)")
        with pytest.raises(CellReferenceError) as exc:
            building.insert_distribution(
                "book.xlsx", "Sheet1", "B1",
                "VoseNormal",
                [{"name": "mu", "value": 0}, {"name": "sigma", "value": 1}],
                dry_run=False,
            )
        assert "non-Vose" in str(exc.value)
        # And the existing cell wasn't touched.
        assert (
            fake_excel.cells[("book.xlsx", "Sheet1", "B1")].formula
            == "=SUM(A1:A10)"
        )

    def test_allows_overwriting_existing_vose_formula(
        self,
        fake_excel: WritableFakeExcel,
        bridge_with_audit: tuple[ModelRiskBridge, Path],
    ) -> None:
        fake_excel.write_cell(
            "book.xlsx", "Sheet1", "B1", "=VoseModPERT(1,2,3)"
        )
        result = building.insert_distribution(
            "book.xlsx", "Sheet1", "B1",
            "VoseNormal",
            [{"name": "mu", "value": 0}, {"name": "sigma", "value": 1}],
            dry_run=False,
        )
        assert result.written is True
        assert (
            fake_excel.cells[("book.xlsx", "Sheet1", "B1")].formula
            == "=VoseNormal(0,1)"
        )

    def test_replace_constant_does_overwrite_non_vose(
        self,
        fake_excel: WritableFakeExcel,
        bridge_with_audit: tuple[ModelRiskBridge, Path],
    ) -> None:
        """The one tool explicitly allowed to overwrite a non-Vose cell."""
        # A1 starts with no formula but a plain number (42).
        fake_excel.cells[("book.xlsx", "Sheet1", "A1")] = CellInfo(
            ref=CellRef(workbook="book.xlsx", sheet="Sheet1", cell="A1"),
            formula="", value=42, cell_type="number",
        )
        result = building.replace_constant_with_distribution(
            "book.xlsx", "Sheet1", "A1",
            "VoseNormal",
            [{"name": "mu", "value": 42}, {"name": "sigma", "value": 4.2}],
            input_name="Cost",
            dry_run=False,
        )
        assert result.written is True
        new_formula = fake_excel.cells[("book.xlsx", "Sheet1", "A1")].formula
        assert new_formula == '=VoseInput("Cost")+VoseNormal(42,4.2)'


# ----------------------------------------------------------------------
# Unknown function — UnknownFunctionError with suggestion
# ----------------------------------------------------------------------


class TestUnknownFunction:
    def test_insert_distribution_unknown(
        self,
        bridge_with_audit: tuple[ModelRiskBridge, Path],
    ) -> None:
        with pytest.raises(UnknownFunctionError) as exc:
            building.insert_distribution(
                "book.xlsx", "Sheet1", "B1",
                "VoseFoo",
                [{"name": "mu", "value": 0}],
            )
        assert "Did you mean" in str(exc.value)

    def test_fit_unknown_family(
        self,
        bridge_with_audit: tuple[ModelRiskBridge, Path],
    ) -> None:
        with pytest.raises(UnknownFunctionError):
            building.fit_distribution_to_data(
                "book.xlsx", "Sheet1", "B1",
                data_range="A1:A100",
                family="BogusDistribution",
            )


# ----------------------------------------------------------------------
# Restore round-trip (spec §13 Phase 3 acceptance)
# ----------------------------------------------------------------------


class TestRestoreRoundTrip:
    def test_restore_returns_pre_state(
        self,
        fake_excel: WritableFakeExcel,
        bridge_with_audit: tuple[ModelRiskBridge, Path],
    ) -> None:
        # Start with a known Vose formula.
        fake_excel.write_cell(
            "book.xlsx", "Sheet1", "B1", "=VoseModPERT(1,2,3)"
        )
        # Overwrite via the building tool.
        building.insert_distribution(
            "book.xlsx", "Sheet1", "B1",
            "VoseNormal",
            [{"name": "mu", "value": 0}, {"name": "sigma", "value": 1}],
            dry_run=False,
        )
        assert (
            fake_excel.cells[("book.xlsx", "Sheet1", "B1")].formula
            == "=VoseNormal(0,1)"
        )
        # Restore.
        result = restore.restore_cell("book.xlsx", "Sheet1", "B1")
        assert result.written is True
        assert result.formula == "=VoseModPERT(1,2,3)"
        assert (
            fake_excel.cells[("book.xlsx", "Sheet1", "B1")].formula
            == "=VoseModPERT(1,2,3)"
        )

    def test_restore_missing_entry_raises(
        self,
        fake_excel: WritableFakeExcel,
        bridge_with_audit: tuple[ModelRiskBridge, Path],
    ) -> None:
        with pytest.raises(CellReferenceError):
            restore.restore_cell("book.xlsx", "Sheet1", "Z99")

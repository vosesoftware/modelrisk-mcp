"""Workbook-state schemas: cell/range/workbook references and Vose-tagged cells."""

from __future__ import annotations

import re
from typing import Any

from pydantic import BaseModel, ConfigDict, Field, field_validator

# Excel cell reference: column letters + row number, e.g. "B12", "AA1".
_CELL_RE = re.compile(r"^[A-Z]{1,3}[1-9]\d{0,6}$")

# Range reference: two cell refs separated by ":", or a single cell ref.
_RANGE_RE = re.compile(r"^[A-Z]{1,3}[1-9]\d{0,6}(:[A-Z]{1,3}[1-9]\d{0,6})?$")

# Sheet name normalisation: anything between an optional "SheetName!" prefix
# and the cell/range body.
_SHEET_CELL_SPLIT_RE = re.compile(r"^(?:(?P<sheet>[^!]+)!)?(?P<body>.+)$")


def _normalise_cell(raw: str) -> str:
    return raw.strip().upper()


class CellRef(BaseModel):
    """A fully-qualified cell reference: workbook + sheet + cell."""

    workbook: str
    sheet: str
    cell: str

    @field_validator("cell", mode="before")
    @classmethod
    def _normalise(cls, v: Any) -> Any:
        if isinstance(v, str):
            return _normalise_cell(v)
        return v

    @field_validator("cell")
    @classmethod
    def _validate_cell(cls, v: str) -> str:
        if not _CELL_RE.match(v):
            raise ValueError(
                f"Invalid cell reference {v!r}; expected like 'B12' or 'AA1'."
            )
        return v

    @classmethod
    def parse(cls, *, workbook: str, default_sheet: str, ref: str) -> CellRef:
        """Parse a possibly-qualified ref like 'Sheet1!B12' or just 'B12'."""
        m = _SHEET_CELL_SPLIT_RE.match(ref.strip())
        if m is None:
            raise ValueError(f"Could not parse cell reference {ref!r}.")
        sheet = m.group("sheet") or default_sheet
        body = m.group("body")
        return cls(workbook=workbook, sheet=sheet, cell=body)

    @property
    def a1(self) -> str:
        """Render as 'Sheet1!B12'."""
        return f"{self.sheet}!{self.cell}"


class WorkbookInfo(BaseModel):
    name: str
    path: str
    sheets: list[str] = Field(default_factory=list)
    active_sheet: str | None = None


class CellInfo(BaseModel):
    """Snapshot of a single cell as read from Excel."""

    ref: CellRef
    formula: str = ""
    value: float | str | bool | None = None
    number_format: str = ""
    cell_type: str = "general"  # general | formula | number | text | error | empty


class RangeInfo(BaseModel):
    """Snapshot of a contiguous range."""

    model_config = ConfigDict(arbitrary_types_allowed=False)

    workbook: str
    sheet: str
    range_ref: str
    values: list[list[Any]] = Field(default_factory=list)
    formulas: list[list[str]] = Field(default_factory=list)

    @field_validator("range_ref")
    @classmethod
    def _validate_range(cls, v: str) -> str:
        v = v.strip().upper()
        if not _RANGE_RE.match(v):
            raise ValueError(
                f"Invalid range reference {v!r}; expected like 'A1:B12'."
            )
        return v


class ModelRiskInput(BaseModel):
    """A cell wrapped with VoseInput(...)."""

    ref: CellRef
    name: str
    formula: str
    current_value: float | str | None = None


class ModelRiskOutput(BaseModel):
    """A cell wrapped with VoseOutput(...)."""

    ref: CellRef
    name: str
    formula: str
    current_value: float | str | None = None


class DistributionCell(BaseModel):
    """A cell containing a Vose distribution function call."""

    ref: CellRef
    function_name: str  # e.g. "VoseModPERT"
    parameters: list[str] = Field(default_factory=list)  # raw argument strings
    has_input_wrapper: bool = False
    has_output_wrapper: bool = False
    formula: str = ""

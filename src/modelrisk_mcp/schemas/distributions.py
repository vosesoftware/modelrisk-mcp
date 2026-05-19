"""Schemas for distribution insertion and wrapper operations."""

from __future__ import annotations

from typing import Annotated, Any

from pydantic import BaseModel, Field, model_validator

from modelrisk_mcp.schemas.workbook import CellRef

ParamValue = Annotated[
    float | int | bool | str | list[Any],
    Field(
        description=(
            "Parameter value. Number/bool become Excel literals. String is "
            "used verbatim (cell ref like 'B12', range like 'A1:A10', or a "
            "formula fragment without leading '='). List becomes an Excel "
            "array literal like {1,2,3}."
        )
    ),
]


class DistributionParameter(BaseModel):
    """One named argument to a Vose distribution function."""

    name: str
    value: ParamValue


class InsertDistributionRequest(BaseModel):
    """Tool input for `insert_distribution`."""

    cell: CellRef
    function_name: str
    parameters: list[DistributionParameter] = Field(default_factory=list)
    dry_run: bool = True

    @model_validator(mode="after")
    def _names_unique(self) -> InsertDistributionRequest:
        seen: set[str] = set()
        for p in self.parameters:
            if p.name in seen:
                raise ValueError(f"Duplicate parameter name {p.name!r}.")
            seen.add(p.name)
        return self


class InsertResult(BaseModel):
    """Tool output for `insert_distribution` and friends."""

    cell: CellRef
    formula: str
    written: bool  # True if dry_run was False and the write succeeded
    previous_formula: str | None = None  # populated when overwriting


class WrapResult(BaseModel):
    """Tool output for `wrap_with_input` / `wrap_with_output`."""

    cell: CellRef
    formula: str
    written: bool
    previous_formula: str | None = None

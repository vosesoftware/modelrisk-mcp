"""Schemas for simulation results and audit findings."""

from __future__ import annotations

from typing import Literal

from pydantic import BaseModel, Field

from modelrisk_mcp.schemas.workbook import CellRef


class SimulationResult(BaseModel):
    output_name: str
    iterations: int = Field(ge=0)
    mean: float
    stdev: float
    variance: float | None = None
    skewness: float | None = None
    kurtosis: float | None = None
    min: float
    max: float
    percentiles: dict[float, float] = Field(default_factory=dict)


class CorrelationEntry(BaseModel):
    name_a: str
    name_b: str
    pearson: float | None = None
    spearman: float | None = None


class CorrelationMatrix(BaseModel):
    """Result of `get_correlation_matrix`. `names` is the ordered list of
    variables; `pearson[i][j]` and `spearman[i][j]` give the two
    correlation matrices. NaN slots become None."""

    names: list[str] = Field(default_factory=list)
    pearson: list[list[float | None]] = Field(default_factory=list)
    spearman: list[list[float | None]] = Field(default_factory=list)
    iterations: int = 0


class SensitivityEntry(BaseModel):
    input_name: str
    correlation: float
    regression_coefficient: float | None = None


class SensitivityRanking(BaseModel):
    """Tornado data for one output. Entries sorted by abs(correlation) desc."""

    output_name: str
    entries: list[SensitivityEntry] = Field(default_factory=list)
    iterations: int = 0


class AuditFinding(BaseModel):
    severity: Literal["error", "warning", "info"]
    cell: CellRef | None = None
    rule_id: str
    message: str
    suggested_fix: str | None = None


class AuditReport(BaseModel):
    findings: list[AuditFinding] = Field(default_factory=list)

    @property
    def errors(self) -> list[AuditFinding]:
        return [f for f in self.findings if f.severity == "error"]

    @property
    def warnings(self) -> list[AuditFinding]:
        return [f for f in self.findings if f.severity == "warning"]

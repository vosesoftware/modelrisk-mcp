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


class SensitivityEntry(BaseModel):
    input_name: str
    correlation: float
    regression_coefficient: float | None = None


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

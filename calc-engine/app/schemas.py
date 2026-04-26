from typing import Any

from pydantic import BaseModel, Field


class PayrollSource(BaseModel):
    sheet: str
    headerRange: str
    dataRange: str
    filterColumn: str


class PayrollMetrics(BaseModel):
    totalRows: int = Field(ge=0)
    includedRows: int = Field(ge=0)
    loadTimeMs: float = Field(ge=0)


class ModelPeriod(BaseModel):
    date: str
    label: str
    financialYear: int


class ModelConfig(BaseModel):
    lastActualsDate: str
    modelEndDate: str
    calculationStartDate: str
    calculationEndDate: str
    calculationMonths: int = Field(ge=1)
    financialYearEndMonth: int = Field(ge=1, le=12)
    periods: list[ModelPeriod]


class PayrollLoadPreviewRequest(BaseModel):
    section: str
    model: ModelConfig
    source: PayrollSource
    assumptions: dict[str, Any]
    metrics: PayrollMetrics
    headers: list[str]
    rows: list[dict[str, Any]]

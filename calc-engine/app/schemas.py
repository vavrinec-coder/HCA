from typing import Any

from pydantic import BaseModel, Field


class PayrollSource(BaseModel):
    sheet: str
    headerRange: str
    dataRange: str
    filterColumn: str
    storeFilterColumn: str | None = None


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
    userKey: str | None = None
    model: ModelConfig
    source: PayrollSource
    assumptions: dict[str, Any]
    metrics: PayrollMetrics
    headers: list[str]
    rows: list[dict[str, Any]]


class PayrollLoadDetailRequest(BaseModel):
    userKey: str | None = None
    outputKey: str
    periodEndDate: str
    unitId: str


class PayrollLoadDetailBatchItem(BaseModel):
    outputKey: str
    periodEndDate: str
    unitId: str


class PayrollLoadDetailBatchRequest(BaseModel):
    userKey: str | None = None
    items: list[PayrollLoadDetailBatchItem] = Field(min_length=1, max_length=1000)


class PayrollLoadDetailBatchResponse(BaseModel):
    status: str
    values: list[float]
    foundCount: int = 0
    reason: str | None = None


class ClientLogRequest(BaseModel):
    source: str
    stage: str
    level: str = "error"
    message: str
    context: dict[str, Any] = Field(default_factory=dict)

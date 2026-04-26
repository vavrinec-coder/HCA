import os
from typing import Any

from fastapi import FastAPI
from fastapi.middleware.cors import CORSMiddleware
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


class PayrollLoadPreviewRequest(BaseModel):
    section: str
    source: PayrollSource
    metrics: PayrollMetrics
    headers: list[str]
    rows: list[dict[str, Any]]


def _cors_origins() -> list[str]:
    raw = os.getenv(
        "CORS_ORIGINS",
        "https://localhost:3000,https://127.0.0.1:3000",
    )
    return [origin.strip() for origin in raw.split(",") if origin.strip()]


app = FastAPI(title="XF1 External Calc Engine", version="0.1.0")

app.add_middleware(
    CORSMiddleware,
    allow_origins=_cors_origins(),
    allow_credentials=False,
    allow_methods=["GET", "POST", "OPTIONS"],
    allow_headers=["*"],
)


@app.get("/health")
def health() -> dict[str, str]:
    return {"status": "ok"}


@app.post("/payroll/load-preview")
def payroll_load_preview(payload: PayrollLoadPreviewRequest) -> dict[str, Any]:
    sample_keys = list(payload.rows[0].keys()) if payload.rows else []

    return {
        "status": "received",
        "section": payload.section,
        "source": payload.source.model_dump(),
        "totalRows": payload.metrics.totalRows,
        "includedRows": payload.metrics.includedRows,
        "headersCount": len(payload.headers),
        "receivedRows": len(payload.rows),
        "loadTimeMs": payload.metrics.loadTimeMs,
        "sampleKeys": sample_keys[:10],
    }

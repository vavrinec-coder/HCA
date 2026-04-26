import os
from typing import Any

from fastapi import FastAPI
from fastapi.middleware.cors import CORSMiddleware

from app.payroll_headcount import calculate_headcount
from app.schemas import PayrollLoadPreviewRequest


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
    headcount = calculate_headcount(payload.headers, payload.rows, payload.model)

    return {
        "status": "received",
        "section": payload.section,
        "model": {
            "lastActualsDate": payload.model.lastActualsDate,
            "modelEndDate": payload.model.modelEndDate,
            "calculationStartDate": payload.model.calculationStartDate,
            "calculationEndDate": payload.model.calculationEndDate,
            "calculationMonths": payload.model.calculationMonths,
            "financialYearEndMonth": payload.model.financialYearEndMonth,
            "firstPeriod": payload.model.periods[0].model_dump() if payload.model.periods else None,
            "lastPeriod": payload.model.periods[-1].model_dump() if payload.model.periods else None,
        },
        "source": payload.source.model_dump(),
        "totalRows": payload.metrics.totalRows,
        "includedRows": payload.metrics.includedRows,
        "headersCount": len(payload.headers),
        "receivedRows": len(payload.rows),
        "loadTimeMs": payload.metrics.loadTimeMs,
        "sampleKeys": sample_keys[:10],
        "outputs": {
            "headcount": headcount,
        },
    }

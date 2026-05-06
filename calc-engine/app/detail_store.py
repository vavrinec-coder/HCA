import os
from typing import Any
from uuid import uuid4

try:
    import psycopg
except ImportError:  # pragma: no cover - exercised in deployed environment
    psycopg = None


SCHEMA_SQL = [
    """
    CREATE TABLE IF NOT EXISTS calc_runs (
    run_id TEXT PRIMARY KEY,
    user_key TEXT NOT NULL,
    created_at TIMESTAMPTZ NOT NULL DEFAULT NOW(),
    model_last_actuals_date DATE,
    model_end_date DATE,
    calculation_months INTEGER,
    included_rows INTEGER,
    status TEXT NOT NULL
)
    """,
    """
    CREATE TABLE IF NOT EXISTS calcs_detail_outputs (
    run_id TEXT NOT NULL REFERENCES calc_runs(run_id) ON DELETE CASCADE,
    user_key TEXT NOT NULL,
    unit_id TEXT NOT NULL,
    department TEXT,
    period_end_date DATE NOT NULL,
    output_key TEXT NOT NULL,
    value NUMERIC NOT NULL,
    PRIMARY KEY (run_id, unit_id, period_end_date, output_key)
)
    """,
    """
    CREATE INDEX IF NOT EXISTS idx_calcs_detail_outputs_lookup
    ON calcs_detail_outputs (user_key, output_key, period_end_date, unit_id)
    """,
]


def save_latest_run(
    user_key: str | None,
    payload: Any,
    detail_rows: list[dict[str, Any]],
    database_url: str | None = None,
) -> dict[str, Any]:
    clean_user_key = normalize_user_key(user_key)
    if not clean_user_key:
        return {"status": "skipped", "reason": "missing_user_key", "rowsSaved": 0}

    database_url = database_url or os.getenv("DATABASE_URL")
    if not database_url:
        return {
            "status": "skipped",
            "reason": "database_not_configured",
            "rowsSaved": 0,
        }

    if psycopg is None:
        return {
            "status": "skipped",
            "reason": "database_driver_not_installed",
            "rowsSaved": 0,
        }

    run_id = str(uuid4())

    with psycopg.connect(database_url) as connection:
        with connection.transaction():
            for statement in SCHEMA_SQL:
                connection.execute(statement)
            connection.execute(
                "DELETE FROM calc_runs WHERE user_key = %s",
                (clean_user_key,),
            )
            connection.execute(
                """
                INSERT INTO calc_runs (
                    run_id,
                    user_key,
                    model_last_actuals_date,
                    model_end_date,
                    calculation_months,
                    included_rows,
                    status
                )
                VALUES (%s, %s, %s, %s, %s, %s, %s)
                """,
                (
                    run_id,
                    clean_user_key,
                    payload.model.lastActualsDate,
                    payload.model.modelEndDate,
                    payload.model.calculationMonths,
                    payload.metrics.includedRows,
                    "latest",
                ),
            )
            if detail_rows:
                with connection.cursor().copy(
                    """
                    COPY calcs_detail_outputs (
                        run_id,
                        user_key,
                        unit_id,
                        department,
                        period_end_date,
                        output_key,
                        value
                    )
                    FROM STDIN
                    """
                ) as copy:
                    for record in iter_detail_records(
                        run_id,
                        clean_user_key,
                        detail_rows,
                    ):
                        copy.write_row(record)

    return {"status": "saved", "runId": run_id, "rowsSaved": len(detail_rows)}


def build_detail_records(
    run_id: str,
    user_key: str,
    detail_rows: list[dict[str, Any]],
) -> list[tuple[Any, ...]]:
    return list(iter_detail_records(run_id, user_key, detail_rows))


def iter_detail_records(
    run_id: str,
    user_key: str,
    detail_rows: list[dict[str, Any]],
) -> Any:
    for row in detail_rows:
        yield (
            run_id,
            user_key,
            row["unit_id"],
            row.get("department"),
            row["period_end_date"],
            row["output_key"],
            row["value"],
        )


def normalize_user_key(value: str | None) -> str:
    return str(value or "").strip().lower()

import json
import os
from contextlib import contextmanager
from typing import Any
from uuid import uuid4

try:
    import psycopg
except ImportError:  # pragma: no cover - exercised in deployed environment
    psycopg = None

try:
    from psycopg_pool import ConnectionPool
except ImportError:  # pragma: no cover - exercised in deployed environment
    ConnectionPool = None


_pool: Any | None = None
_pool_database_url: str | None = None
_schema_initialized = False


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


def get_database_url(database_url: str | None = None) -> str | None:
    return database_url or os.getenv("DATABASE_URL")


def detail_store_unavailable_reason(
    user_key: str | None,
    database_url: str | None,
) -> str | None:
    if not normalize_user_key(user_key):
        return "missing_user_key"
    if not database_url:
        return "database_not_configured"
    if psycopg is None:
        return "database_driver_not_installed"
    if ConnectionPool is None:
        return "database_pool_not_installed"
    return None


def ensure_schema(connection: Any) -> None:
    for statement in SCHEMA_SQL:
        connection.execute(statement)


def initialize_detail_store(database_url: str | None = None) -> dict[str, Any]:
    global _pool, _pool_database_url, _schema_initialized

    database_url = get_database_url(database_url)
    if not database_url:
        return {"status": "skipped", "reason": "database_not_configured"}

    if psycopg is None:
        return {"status": "skipped", "reason": "database_driver_not_installed"}

    if ConnectionPool is None:
        return {"status": "skipped", "reason": "database_pool_not_installed"}

    if _pool is not None and _pool_database_url == database_url:
        return {"status": "ready"}

    if _pool is not None:
        _pool.close()

    _pool = ConnectionPool(
        conninfo=database_url,
        min_size=1,
        max_size=int(os.getenv("DETAIL_DB_POOL_MAX_SIZE", "5")),
        open=True,
    )
    _pool_database_url = database_url
    _schema_initialized = False

    with _pool.connection() as connection:
        ensure_schema(connection)

    _schema_initialized = True
    return {"status": "ready"}


def close_detail_store() -> None:
    global _pool, _pool_database_url, _schema_initialized

    if _pool is not None:
        _pool.close()

    _pool = None
    _pool_database_url = None
    _schema_initialized = False


@contextmanager
def detail_connection(database_url: str | None = None):
    database_url = get_database_url(database_url)

    if _pool is not None and (_pool_database_url == database_url or database_url is None):
        with _pool.connection() as connection:
            yield connection
        return

    if not database_url:
        raise RuntimeError("database_not_configured")

    with psycopg.connect(database_url) as connection:
        yield connection


def save_latest_run(
    user_key: str | None,
    payload: Any,
    detail_rows: list[dict[str, Any]],
    database_url: str | None = None,
) -> dict[str, Any]:
    detail_rows_to_save = nonzero_detail_rows(detail_rows)
    clean_user_key = normalize_user_key(user_key)
    database_url = get_database_url(database_url)
    unavailable_reason = detail_store_unavailable_reason(clean_user_key, database_url)
    if unavailable_reason is not None:
        return {
            "status": "skipped",
            "reason": unavailable_reason,
            "rowsPrepared": len(detail_rows),
            "rowsSaved": 0,
        }

    run_id = str(uuid4())

    with detail_connection(database_url) as connection:
        with connection.transaction():
            ensure_schema(connection)
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
            if detail_rows_to_save:
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
                        detail_rows_to_save,
                    ):
                        copy.write_row(record)

    return {
        "status": "saved",
        "runId": run_id,
        "rowsPrepared": len(detail_rows),
        "rowsSaved": len(detail_rows_to_save),
    }


def load_detail_value(
    user_key: str | None,
    output_key: str,
    period_end_date: str,
    unit_id: str,
    database_url: str | None = None,
) -> dict[str, Any]:
    clean_user_key = normalize_user_key(user_key)
    clean_output_key = str(output_key or "").strip()
    clean_period_end_date = str(period_end_date or "").strip()
    clean_unit_id = str(unit_id or "").strip()
    database_url = get_database_url(database_url)
    unavailable_reason = detail_store_unavailable_reason(clean_user_key, database_url)
    if unavailable_reason is not None:
        return {"status": "skipped", "reason": unavailable_reason, "value": 0}

    with detail_connection(database_url) as connection:
        row = connection.execute(
            """
            SELECT value
            FROM calcs_detail_outputs
            WHERE user_key = %s
              AND output_key = %s
              AND period_end_date = %s
              AND unit_id = %s
            LIMIT 1
            """,
            (
                clean_user_key,
                clean_output_key,
                clean_period_end_date,
                clean_unit_id,
            ),
        ).fetchone()

    if row is None:
        return {"status": "not_found", "value": 0}

    return {"status": "found", "value": float(row[0])}


def load_detail_values(
    user_key: str | None,
    items: list[Any],
    database_url: str | None = None,
) -> dict[str, Any]:
    clean_user_key = normalize_user_key(user_key)
    database_url = get_database_url(database_url)

    unavailable_reason = detail_store_unavailable_reason(clean_user_key, database_url)
    if unavailable_reason is not None:
        return {
            "status": "skipped",
            "reason": unavailable_reason,
            "values": [0.0] * len(items),
            "foundCount": 0,
        }

    normalized_items = [normalize_detail_item(item) for item in items]
    values = [0.0] * len(normalized_items)

    if not normalized_items:
        return {"status": "ok", "values": [], "foundCount": 0}

    payload_json = json.dumps(
        [
            {
                "ord": index,
                "output_key": item["output_key"],
                "period_end_date": item["period_end_date"],
                "unit_id": item["unit_id"],
            }
            for index, item in enumerate(normalized_items)
        ]
    )

    with detail_connection(database_url) as connection:
        rows = connection.execute(
            """
            WITH input AS (
                SELECT *
                FROM jsonb_to_recordset(%s::jsonb)
                AS x(
                    ord integer,
                    output_key text,
                    period_end_date date,
                    unit_id text
                )
            )
            SELECT input.ord, detail.value
            FROM input
            LEFT JOIN calcs_detail_outputs AS detail
              ON detail.user_key = %s
             AND detail.output_key = input.output_key
             AND detail.period_end_date = input.period_end_date
             AND detail.unit_id = input.unit_id
            ORDER BY input.ord
            """,
            (payload_json, clean_user_key),
        ).fetchall()

    found_count = 0
    for ordinal, value in rows:
        if value is not None:
            values[int(ordinal)] = float(value)
            found_count += 1

    return {
        "status": "ok",
        "values": values,
        "foundCount": found_count,
    }


def read_item_value(item: Any, attr_name: str, key_name: str) -> Any:
    if hasattr(item, attr_name):
        return getattr(item, attr_name)
    if isinstance(item, dict):
        return item.get(key_name)
    return None


def normalize_detail_item(item: Any) -> dict[str, str]:
    return {
        "output_key": str(
            read_item_value(item, "outputKey", "outputKey") or ""
        ).strip(),
        "period_end_date": str(
            read_item_value(item, "periodEndDate", "periodEndDate") or ""
        ).strip(),
        "unit_id": str(read_item_value(item, "unitId", "unitId") or "").strip(),
    }


def nonzero_detail_rows(detail_rows: list[dict[str, Any]]) -> list[dict[str, Any]]:
    return [row for row in detail_rows if float(row.get("value") or 0) != 0.0]


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

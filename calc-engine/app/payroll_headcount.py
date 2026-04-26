from calendar import monthrange
from collections import defaultdict
from datetime import date, datetime, timedelta
from typing import Any

from app.schemas import ModelConfig


DEPARTMENT_FIELD_INDEX = 3
START_DATE_FIELD_INDEX = 4
END_DATE_FIELD_INDEX = 5
FS_CATEGORY_FIELD_INDEX = 1
STATUS_FIELD_INDEX = 2


def calculate_payroll_outputs(
    headers: list[str],
    rows: list[dict[str, Any]],
    model: ModelConfig,
) -> dict[str, Any]:
    department_field = field_name(headers, DEPARTMENT_FIELD_INDEX, "department")
    start_date_field = field_name(headers, START_DATE_FIELD_INDEX, "start date")
    end_date_field = field_name(headers, END_DATE_FIELD_INDEX, "end date")
    fs_category_field = field_name(headers, FS_CATEGORY_FIELD_INDEX, "FS category")
    status_field = field_name(headers, STATUS_FIELD_INDEX, "status")
    periods = [
        {
            "date": parse_iso_date(period.date),
            "label": period.label,
            "financialYear": period.financialYear,
        }
        for period in model.periods
    ]
    salary_fields = salary_field_by_period(headers, periods)
    headcount_totals: dict[str, list[float]] = defaultdict(
        lambda: [0.0] * len(periods)
    )
    base_salary_totals: dict[str, list[float]] = defaultdict(
        lambda: [0.0] * len(periods)
    )
    domestic_salary_totals: dict[str, list[float]] = defaultdict(
        lambda: [0.0] * len(periods)
    )
    international_salary_totals: dict[str, list[float]] = defaultdict(
        lambda: [0.0] * len(periods)
    )
    cogs_salary_totals: dict[str, list[float]] = defaultdict(
        lambda: [0.0] * len(periods)
    )
    skipped_rows = 0

    for row in rows:
        department = normalize_text(row.get(department_field))
        start_date = parse_date_value(row.get(start_date_field))
        end_date = parse_date_value(row.get(end_date_field)) or parse_iso_date(
            model.calculationEndDate
        )

        if not department or start_date is None:
            skipped_rows += 1
            continue

        status = normalize_key(row.get(status_field))
        fs_category = normalize_key(row.get(fs_category_field))

        for index, period in enumerate(periods):
            month_end = period["date"]
            month_start = date(month_end.year, month_end.month, 1)
            fte = monthly_fte(start_date, end_date, month_start, month_end)
            annual_salary = parse_number(row.get(salary_fields[index]))
            monthly_salary = (annual_salary / 12) * fte
            headcount_totals[department][index] += fte
            base_salary_totals[department][index] += monthly_salary

            if status == "domestic":
                domestic_salary_totals[department][index] += monthly_salary
            elif status == "international":
                international_salary_totals[department][index] += monthly_salary

            if fs_category == "cos":
                cogs_salary_totals[department][index] += monthly_salary

    departments = sorted(headcount_totals)
    domestic_departments = sorted(domestic_salary_totals)
    international_departments = sorted(international_salary_totals)
    cogs_departments = sorted(cogs_salary_totals)

    return {
        "headcount": {
            "table": output_table(departments, periods, headcount_totals, decimals=2),
            "departments": departments,
            "periods": serialize_periods(periods),
            "skippedRows": skipped_rows,
            "fieldMap": {
                "department": department_field,
                "startDate": start_date_field,
                "endDate": end_date_field,
            },
        },
        "baseSalary": {
            "total": {
                "table": output_table(
                    departments, periods, base_salary_totals, decimals=0
                ),
                "departments": departments,
            },
            "domestic": {
                "table": output_table(
                    domestic_departments,
                    periods,
                    domestic_salary_totals,
                    decimals=0,
                ),
                "departments": domestic_departments,
            },
            "international": {
                "table": output_table(
                    international_departments,
                    periods,
                    international_salary_totals,
                    decimals=0,
                ),
                "departments": international_departments,
            },
            "cogs": {
                "table": output_table(
                    cogs_departments, periods, cogs_salary_totals, decimals=0
                ),
                "departments": cogs_departments,
            },
            "periods": serialize_periods(periods),
            "salaryFieldByPeriod": salary_fields,
        },
    }


def monthly_fte(
    employee_start: date,
    employee_end: date,
    month_start: date,
    month_end: date,
) -> float:
    active_start = max(employee_start, month_start)
    active_end = min(employee_end, month_end)

    if active_start > active_end:
        return 0.0

    active_days = (active_end - active_start).days + 1
    days_in_month = monthrange(month_end.year, month_end.month)[1]
    return active_days / days_in_month


def salary_field_by_period(
    headers: list[str],
    periods: list[dict[str, Any]],
) -> list[str | None]:
    available_salary_fields = {str(header).strip(): header for header in headers}
    return [
        available_salary_fields.get(str(period["financialYear"]))
        for period in periods
    ]


def output_table(
    departments: list[str],
    periods: list[dict[str, Any]],
    totals: dict[str, list[float]],
    decimals: int,
) -> list[list[Any]]:
    table = [["Department", *[period["label"] for period in periods]]]
    for department in departments:
        table.append(
            [
                department,
                *[round(value, decimals) for value in totals[department]],
            ]
        )
    return table


def serialize_periods(periods: list[dict[str, Any]]) -> list[dict[str, Any]]:
    return [
        {
            "date": period["date"].isoformat(),
            "label": period["label"],
            "financialYear": period["financialYear"],
        }
        for period in periods
    ]


def field_name(headers: list[str], index: int, fallback: str) -> str:
    if index >= len(headers):
        raise ValueError(f"Payroll data is missing expected field: {fallback}")
    return headers[index]


def normalize_text(value: Any) -> str:
    return str(value or "").strip()


def normalize_key(value: Any) -> str:
    return normalize_text(value).lower()


def parse_number(value: Any) -> float:
    if value in (None, "", "-"):
        return 0.0

    if isinstance(value, (int, float)):
        return float(value)

    text = str(value).strip().replace(",", "")
    if not text or text == "-":
        return 0.0

    return float(text)


def parse_date_value(value: Any) -> date | None:
    if value in (None, ""):
        return None

    if isinstance(value, (int, float)):
        return date(1899, 12, 30) + timedelta(days=int(value))

    if isinstance(value, date):
        return value

    text = str(value).strip()
    if not text:
        return None

    for fmt in ("%Y-%m-%d", "%d-%b-%y", "%d-%b-%Y", "%d/%b/%y", "%d/%b/%Y"):
        try:
            return datetime.strptime(text, fmt).date()
        except ValueError:
            pass

    raise ValueError(f"Invalid date value: {text}")


def parse_iso_date(value: str) -> date:
    return datetime.strptime(value, "%Y-%m-%d").date()

# Payroll Load Preview Payload

The add-in sends this shape to `POST /payroll/load-preview`:

```json
{
  "section": "Payroll",
  "model": {
    "lastActualsDate": "2026-03-31",
    "modelEndDate": "2028-04-30",
    "calculationStartDate": "2026-04-30",
    "calculationEndDate": "2028-04-30",
    "calculationMonths": 25,
    "financialYearEndMonth": 4,
    "periods": [
      {
        "date": "2026-04-30",
        "label": "Apr 2026",
        "financialYear": 2026
      }
    ]
  },
  "source": {
    "sheet": "PayrollData",
    "headerRange": "B4:R4",
    "dataRange": "B5:R1531",
    "filterColumn": "R"
  },
  "output": {
    "sheet": "HCA_Output",
    "headcountStartCell": "E4",
    "baseSalaryTotalStartCell": "E17",
    "baseSalaryDomesticStartCell": "E30",
    "baseSalaryInternationalStartCell": "E44",
    "baseSalaryCogsStartCell": "E57"
  },
  "metrics": {
    "totalRows": 1527,
    "includedRows": 123,
    "loadTimeMs": 250
  },
  "headers": ["Employee ID", "Name", "Include in Load"],
  "rows": [
    {
      "Employee ID": "E001",
      "Name": "Example",
      "Include in Load": 1
    }
  ]
}
```

The MVP backend returns only a summary. It does not store data or calculate payroll.

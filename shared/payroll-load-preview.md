# Payroll Load Preview Payload

The add-in sends this shape to `POST /payroll/load-preview`:

```json
{
  "section": "Payroll",
  "source": {
    "sheet": "PayrollData",
    "headerRange": "B4:S4",
    "dataRange": "B5:S1531",
    "filterColumn": "S"
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

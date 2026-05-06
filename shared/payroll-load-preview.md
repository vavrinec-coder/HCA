# Payroll Load Preview Payload

The add-in sends this shape to `POST /payroll/load-preview`:

```json
{
  "section": "Payroll",
  "userKey": "vavrinec@xf1advisory.com",
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
    "baseSalaryCogsStartCell": "E57",
    "medicalStartCell": "E70",
    "retirement401kStartCell": "E83",
    "otherBenefitsStartCell": "E96",
    "bonusAccrualStartCell": "E110",
    "bonusPayoutStartCell": "E124"
  },
  "assumptions": {
    "benefits": {
      "medical": {
        "domestic": 2464,
        "international": 2162
      },
      "retirement401k": {
        "domestic": 432,
        "international": 501
      },
      "otherBenefits": {
        "domestic": 157,
        "international": 20
      }
    },
    "bonus": {
      "cap": 2,
      "netNewArrAchieved": [1.2],
      "burnMultipleAchieved": [0.8],
      "executivePlan": {
        "netNewArrWeight": 0.75,
        "burnMultipleWeight": 0.25
      },
      "incentivePlan": {
        "netNewArrWeight": 1,
        "burnMultipleWeight": 0
      }
    }
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

The backend returns a summary plus calculated payroll output tables. It does not store data.
If `DATABASE_URL` is configured, the backend also stores latest-run employee/month/output detail rows for the supplied `userKey`.

The response includes detail save status:

```json
{
  "detailSave": {
    "status": "saved",
    "runId": "example-run-id",
    "rowsSaved": 283750
  }
}
```

If detail storage is not available, Payroll output still succeeds and the response reports:

```json
{
  "detailSave": {
    "status": "skipped",
    "reason": "database_not_configured",
    "rowsSaved": 0
  }
}
```

import os
import unittest

try:
    from app.main import payroll_load_preview
    from app.schemas import PayrollLoadPreviewRequest
except ModuleNotFoundError:
    payroll_load_preview = None
    PayrollLoadPreviewRequest = None


class PayrollApiTests(unittest.TestCase):
    def test_load_preview_skips_detail_storage_when_database_is_not_configured(self):
        if payroll_load_preview is None:
            self.skipTest("FastAPI test dependency is not installed.")

        previous_database_url = os.environ.pop("DATABASE_URL", None)
        self.addCleanup(self._restore_database_url, previous_database_url)

        response = payroll_load_preview(
            PayrollLoadPreviewRequest.model_validate(
                {
                    "section": "Payroll",
                    "userKey": "user@example.com",
                    "model": {
                        "lastActualsDate": "2026-03-31",
                        "modelEndDate": "2026-04-30",
                        "calculationStartDate": "2026-04-30",
                        "calculationEndDate": "2026-04-30",
                        "calculationMonths": 1,
                        "financialYearEndMonth": 4,
                        "periods": [
                            {
                                "date": "2026-04-30",
                                "label": "Apr 2026",
                                "financialYear": 2026,
                            }
                        ],
                    },
                    "source": {
                        "sheet": "PayrollData",
                        "headerRange": "B4:R4",
                        "dataRange": "B5:R1531",
                        "filterColumn": "R",
                    },
                    "assumptions": {},
                    "metrics": {"totalRows": 1, "includedRows": 1, "loadTimeMs": 1},
                    "headers": [
                        "EmployeeID",
                        "FS_Category",
                        "Status",
                        "Department",
                        "Start Date",
                        "Termination Date",
                        "Bonus Plan",
                        "Bonus %",
                        "Bonus $",
                        "Payroll Case",
                        "Severance Pay",
                        "2026",
                    ],
                    "rows": [
                        {
                            "EmployeeID": "E1",
                            "FS_Category": "OpEx",
                            "Status": "Domestic",
                            "Department": "Sales",
                            "Start Date": "2026-01-01",
                            "Termination Date": "2099-12-31",
                            "Bonus Plan": "na",
                            "Bonus %": 0,
                            "Bonus $": 0,
                            "2026": 120000,
                        }
                    ],
                }
            )
        )

        self.assertEqual(response["detailSave"]["status"], "skipped")
        self.assertEqual(response["detailSave"]["reason"], "database_not_configured")
        self.assertGreaterEqual(response["timings"]["calculationMs"], 0)
        self.assertGreaterEqual(response["timings"]["detailSaveMs"], 0)
        self.assertGreaterEqual(response["timings"]["totalBackendMs"], 0)
        self.assertNotIn("detailRows", response["outputs"])

    @staticmethod
    def _restore_database_url(value):
        if value is not None:
            os.environ["DATABASE_URL"] = value


if __name__ == "__main__":
    unittest.main()

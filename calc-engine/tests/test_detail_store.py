import unittest

from app.detail_store import build_detail_records


class DetailStoreTests(unittest.TestCase):
    def test_build_detail_records_adds_run_and_user_fields(self):
        rows = [
            {
                "unit_id": "E1",
                "department": "Sales",
                "period_end_date": "2026-05-31",
                "output_key": "payroll.output.bonus_payout",
                "value": 1100,
            }
        ]

        records = build_detail_records("run-1", "user@example.com", rows)

        self.assertEqual(
            records,
            [
                (
                    "run-1",
                    "user@example.com",
                    "E1",
                    "Sales",
                    "2026-05-31",
                    "payroll.output.bonus_payout",
                    1100,
                )
            ],
        )


if __name__ == "__main__":
    unittest.main()

"""Unit tests for the metric taxonomy and slice reconciliation."""

from __future__ import annotations

import sys
import unittest
from datetime import date
from pathlib import Path


ROOT = Path(__file__).resolve().parents[1]
sys.path.insert(0, str(ROOT))

from customer_cube import calculate_cube  # noqa: E402
from models import CubeConfig, RevenueRecord  # noqa: E402


def record(customer: str, month: str, amount: float, product: str) -> RevenueRecord:
    year, month_number = (int(part) for part in month.split("-"))
    return RevenueRecord(customer, date(year, month_number, 1), amount, product)


class CustomerCubeTests(unittest.TestCase):
    def test_movement_taxonomy_and_retention(self) -> None:
        records = (
            record("A", "2024-01", 100, "P1"),
            record("A", "2024-01", 50, "P2"),
            record("B", "2024-01", 80, "P1"),
            record("C", "2024-01", 40, "P1"),
            record("A", "2024-04", 120, "P1"),
            record("A", "2024-04", 30, "P2"),
            record("A", "2024-04", 10, "P3"),
            record("C", "2024-04", 40, "P1"),
            record("D", "2024-04", 70, "P1"),
        )
        rows = calculate_cube(records, CubeConfig(slices=("overall",)))
        q2 = next(row for row in rows if row.grain == "quarterly" and row.period == "FY2024-Q2")

        self.assertEqual(q2.opening_revenue, 270)
        self.assertEqual(q2.new, 70)
        self.assertEqual(q2.upsell, 20)
        self.assertEqual(q2.downsell, 20)
        self.assertEqual(q2.cross_sell, 10)
        self.assertEqual(q2.churn, 80)
        self.assertEqual(q2.closing_revenue, 270)
        self.assertAlmostEqual(q2.gross_retention or 0, 170 / 270)
        self.assertAlmostEqual(q2.net_retention or 0, 200 / 270)

    def test_product_and_customer_slices_sum_to_overall(self) -> None:
        records = (
            record("A", "2024-01", 100, "P1"),
            record("A", "2024-04", 120, "P1"),
            record("A", "2024-04", 20, "P2"),
            record("B", "2024-01", 50, "P2"),
        )
        rows = calculate_cube(records, CubeConfig())
        q2_rows = [row for row in rows if row.grain == "quarterly" and row.period == "FY2024-Q2"]
        overall = next(row for row in q2_rows if row.slice_type == "overall")
        products = [row for row in q2_rows if row.slice_type == "product"]
        customers = [row for row in q2_rows if row.slice_type == "customer"]

        for field in ("opening_revenue", "new", "upsell", "downsell", "cross_sell", "churn", "closing_revenue"):
            self.assertAlmostEqual(sum(getattr(row, field) for row in products), getattr(overall, field))
            self.assertAlmostEqual(sum(getattr(row, field) for row in customers), getattr(overall, field))

    def test_period_end_uses_the_calendar_period_end_not_latest_source_row(self) -> None:
        records = (
            record("A", "2024-01", 100, "P1"),
            record("B", "2024-03", 1, "P1"),
            record("A", "2024-04", 120, "P1"),
            record("B", "2024-06", 1, "P1"),
        )
        rows = calculate_cube(
            records,
            CubeConfig(amount_mode="period_end", slices=("overall",)),
        )
        q1 = next(row for row in rows if row.grain == "quarterly" and row.period == "FY2024-Q1")
        q2 = next(row for row in rows if row.grain == "quarterly" and row.period == "FY2024-Q2")

        self.assertEqual(q1.closing_revenue, 1)
        self.assertEqual(q2.opening_revenue, 1)
        self.assertEqual(q2.closing_revenue, 1)
        self.assertEqual(q2.new, 0)
        self.assertEqual(q2.churn, 0)


if __name__ == "__main__":
    unittest.main()

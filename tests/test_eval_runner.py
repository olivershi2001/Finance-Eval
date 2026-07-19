"""End-to-end tests for the customer-cube runner and scorer."""

from __future__ import annotations

import sys
import unittest
from dataclasses import replace
from pathlib import Path
from unittest.mock import patch


ROOT = Path(__file__).resolve().parents[1]
sys.path.insert(0, str(ROOT))

from eval_runner import DEFAULT_INPUT, load_config, load_cube_rows, run_eval  # noqa: E402
from models import CandidateResult  # noqa: E402
from scoring import score_cube  # noqa: E402
from xlsx_reader import load_revenue_records  # noqa: E402


class EvalRunnerTests(unittest.TestCase):
    def test_raw_excel_contract_loads(self) -> None:
        records = load_revenue_records(DEFAULT_INPUT)

        self.assertEqual(len(records), 1_000)
        self.assertEqual(
            {record.product for record in records},
            {"Analytics", "Automation", "Platform", "Security", "Support"},
        )
        self.assertEqual(min(record.month for record in records).isoformat(), "2022-01-01")
        self.assertEqual(max(record.month for record in records).isoformat(), "2025-12-01")

    def test_reference_candidate_passes_frozen_gold(self) -> None:
        report = run_eval()

        self.assertTrue(report.passed)
        self.assertEqual(report.score, 100.0)
        self.assertGreater(report.expected_rows, 150)

    def test_scorer_detects_empty_candidate(self) -> None:
        config = load_config()
        expected = load_cube_rows(ROOT / "scenarios" / "customer_cube_expected.json")

        report = score_cube(expected, (), config)

        self.assertFalse(report.passed)
        self.assertLess(report.score, 10.0)
        self.assertEqual(len(report.missing_keys), 20)

    def test_wrong_metric_is_reported(self) -> None:
        config = load_config()
        expected = load_cube_rows(ROOT / "scenarios" / "customer_cube_expected.json")
        broken = (replace(expected[0], new=expected[0].new + 25),) + expected[1:]

        report = score_cube(expected, broken, config)

        self.assertLess(report.score, 100.0)
        self.assertTrue(any(" new:" in mismatch for mismatch in report.mismatches))


if __name__ == "__main__":
    unittest.main()

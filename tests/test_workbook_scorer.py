"""Tests for the generic XLSX-against-XLSX scorer."""

from __future__ import annotations

import sys
import unittest
from datetime import date
from pathlib import Path


ROOT = Path(__file__).resolve().parents[1]
sys.path.insert(0, str(ROOT))

from workbook_scorer import (  # noqa: E402
    CellRange,
    WorkbookScoreConfig,
    score_workbook_cells,
    score_workbooks,
)
from xlsx_reader import WorkbookCell  # noqa: E402


def cell(value: object, formula: str | None = None) -> WorkbookCell:
    return WorkbookCell(value=value, formula=formula)


class WorkbookScorerTests(unittest.TestCase):
    def test_values_allow_numeric_tolerance_and_text_normalization(self) -> None:
        expected = {
            "Summary": {
                "A1": cell(100.0),
                "A2": cell("Net Revenue Retention"),
                "A3": cell(date(2026, 6, 30)),
            }
        }
        candidate = {
            "summary": {
                "A1": cell(100.005),
                "A2": cell("  net   revenue retention "),
                "A3": cell(date(2026, 6, 30)),
            }
        }

        report = score_workbook_cells(expected, candidate, WorkbookScoreConfig())

        self.assertTrue(report.passed)
        self.assertEqual(report.score, 100.0)

    def test_missing_and_wrong_cells_reduce_coverage_and_value_scores(self) -> None:
        expected = {"Summary": {"A1": cell(10), "A2": cell(20)}}
        candidate = {"Summary": {"A1": cell(999)}}

        report = score_workbook_cells(expected, candidate, WorkbookScoreConfig())

        self.assertFalse(report.passed)
        self.assertEqual(report.covered_cells, 1)
        self.assertEqual(report.matched_cells, 0)
        self.assertEqual(len(report.mismatches), 2)

    def test_sheet_mapping_and_ranges_allow_a_different_candidate_tab_name(self) -> None:
        expected = {"Gold": {"A1": cell(1), "B1": cell(999)}}
        candidate = {"Model Output": {"A1": cell(1), "B1": cell(0)}}
        config = WorkbookScoreConfig(
            sheet_map={"Gold": "Model Output"},
            ranges={"Gold": (CellRange(1, 1, 1, 1),)},
        )

        report = score_workbook_cells(expected, candidate, config)

        self.assertEqual(report.score, 100.0)
        self.assertEqual(report.expected_cells, 1)

    def test_formula_mode_can_require_the_exact_gold_formula(self) -> None:
        expected = {
            "Model": {
                "A1": cell(1),
                "B1": cell(2, "=A1*2"),
            }
        }
        candidate = {"Model": {"A1": cell(1), "B1": cell(2)}}

        values_report = score_workbook_cells(
            expected, candidate, WorkbookScoreConfig(formula_mode="values")
        )
        formulas_report = score_workbook_cells(
            expected, candidate, WorkbookScoreConfig(formula_mode="exact")
        )

        self.assertEqual(values_report.score, 100.0)
        self.assertLess(formulas_report.score, 100.0)
        self.assertTrue(any("formula" in item for item in formulas_report.mismatches))

    def test_extra_cells_are_only_penalized_in_strict_mode(self) -> None:
        expected = {"Output": {"A1": cell(1)}}
        candidate = {"Output": {"A1": cell(1), "B1": cell("extra")}}

        default_report = score_workbook_cells(
            expected, candidate, WorkbookScoreConfig()
        )
        strict_report = score_workbook_cells(
            expected,
            candidate,
            WorkbookScoreConfig(penalize_extra_cells=True),
        )

        self.assertEqual(default_report.score, 100.0)
        self.assertEqual(default_report.extra_cells, 1)
        self.assertLess(strict_report.score, 100.0)

    def test_real_reference_workbook_scores_against_itself(self) -> None:
        workbook = (
            ROOT
            / "outputs"
            / "customer_cube_eval"
            / "customer_cube_reference_output.xlsx"
        )

        report = score_workbooks(workbook, workbook, WorkbookScoreConfig())

        self.assertEqual(report.score, 100.0)
        self.assertGreater(report.expected_cells, 1_000)


if __name__ == "__main__":
    unittest.main()

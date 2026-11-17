"""Data isolation and correction-loop regression tests."""
import copy
import hashlib
import io
import json
from pathlib import Path
import tempfile
import threading
import unittest
from unittest.mock import patch

import app
from lab import Lab, fingerprint


class LabTests(unittest.TestCase):
    def setUp(self):
        self.temp = tempfile.TemporaryDirectory()
        self.root = Path(self.temp.name)
        self.lab = Lab(self.root, threading.RLock())
        self.data = (app.ROOT / "fixtures/customer_cube_raw.xlsx").read_bytes()
        self.identity = "a" * 24
        self.key = ["quarterly", "FY2025-Q4", "overall", "All"]
        self.row = dict(zip(("grain", "period", "slice_type", "slice_value"), self.key))
        self.row.update(opening_revenue=100, new=0, upsell=0, downsell=10, cross_sell=0, churn=0,
                        closing_revenue=90, gross_retention=.9, net_retention=.9, ttm_gross_retention=None)
        self.meta = {"id": self.identity, "filename": "synthetic.xlsx", "created_at": "2026-11-17T00:00:00Z",
                     "amount_mode": "period_sum", "source_sha256": hashlib.sha256(self.data).hexdigest(),
                     "status": "complete", "stage": "Ready", "mode": "model", "model": "test-a", "demo": True}
        self.result = {"rows": [self.row], "reference_rows": [{**self.row, "closing_revenue": 95}],
                       "score": {"score": 90, "passed": False}, "differences": [
                           {"row_key": self.key, "metric": "closing_revenue", "expected": 95, "actual": 90}]}
        self.note = {"id": "note-1", "note": "Verify the closing balance against the source", "row_key": self.key,
                     "status": "flagged", "created_at": self.meta["created_at"]}
        self.save_run(self.identity)

    def tearDown(self):
        self.temp.cleanup()

    def save_run(self, identity, meta=None, result=None):
        path = self.root / identity
        path.mkdir(exist_ok=True)
        (path / "source.xlsx").write_bytes(self.data)
        self.lab.write(path / "run.json", {**(meta or self.meta), "id": identity})
        self.lab.write(path / "result.json", result or self.result)
        self.lab.write(path / "feedback.json", [self.note])

    def case(self, **changes):
        return self.lab.create_case({"run_id": self.identity, "note_id": self.note["id"], "metric": "closing_revenue",
                                     "expected": 95, "title": "Closing balance correction", "rationale": "Reviewed source rows", **changes})

    def test_model_comparison_freezes_identical_source_and_distinct_model_ids(self):
        calls = []
        def create(data, filename, **kwargs):
            calls.append((data, kwargs))
            identity = str(len(calls)) * 24
            self.save_run(identity, {**self.meta, "selected_model": kwargs["selected_model"], "status": "queued"})
            return {"id": identity}
        batch = self.lab.create_comparison({"source_run_id": self.identity, "models": ["test-a", "test-b"]}, create)
        self.assertEqual(len(batch["candidates"]), 2)
        self.assertEqual(calls[0][0], calls[1][0])
        self.assertEqual(calls[0][1]["comparison_id"], calls[1][1]["comparison_id"])
        self.assertEqual(batch["status"], "running")
        self.assertIsNone(batch["same_prompt"])

    def test_comparison_rejects_duplicate_models_or_changed_source(self):
        with self.assertRaisesRegex(ValueError, "different"):
            self.lab.create_comparison({"source_run_id": self.identity, "models": ["test-a", "test-a"]}, None)
        (self.root / self.identity / "source.xlsx").write_bytes(b"changed")
        with self.assertRaisesRegex(ValueError, "fingerprint"):
            self.lab.create_comparison({"source_run_id": self.identity, "models": ["test-a", "test-b"]}, None)

    def test_issues_have_separate_append_only_resolution_history(self):
        original = (self.root / self.identity / "result.json").read_bytes()
        issue = next(i for i in self.lab.issues() if i["category"] == "revenue")
        self.lab.update_issue(issue["id"], {"status": "investigating", "note": "Checked source"})
        self.lab.update_issue(issue["id"], {"status": "resolved", "note": "Corrected in a new version"})
        issue = next(i for i in self.lab.issues() if i["id"] == issue["id"])
        self.assertEqual(issue["status"], "resolved")
        self.assertEqual(len(issue["history"]), 2)
        self.assertEqual((self.root / self.identity / "result.json").read_bytes(), original)

    def test_missing_and_duplicate_rows_are_failures_without_numeric_differences(self):
        result = {**self.result, "rows": [self.row, self.row], "reference_rows": [self.row, {**self.row, "period": "FY2025-Q3"}], "differences": []}
        self.save_run(self.identity, result=result)
        categories = [i["category"] for i in self.lab.issues()]
        self.assertIn("coverage", categories)
        self.assertIn("structure", categories)

    def test_reviewer_answer_can_disagree_with_engine_and_detect_future_regression(self):
        case = self.case(expected=90)
        self.assertEqual(case["provenance"], "Reviewer-entered expected value")
        self.assertEqual(self.lab.evaluate_cases(self.identity)[0]["status"], "pass")
        self.save_run("b" * 24, result={**self.result, "rows": [{**self.row, "closing_revenue": 95}]})
        self.assertEqual(self.lab.evaluate_cases("b" * 24)[0]["status"], "fail")
        self.assertEqual(len(self.lab.cases()), 1)
        self.case(expected=90)
        self.assertEqual(len(self.lab.cases()), 1)

    def test_cases_skip_different_data_and_fail_missing_or_duplicate_rows(self):
        self.case()
        self.save_run("b" * 24, meta={**self.meta, "source_sha256": "different"})
        self.assertEqual(self.lab.evaluate_cases("b" * 24)[0]["status"], "skipped")
        self.save_run("b" * 24, result={**self.result, "rows": []})
        self.assertEqual(self.lab.evaluate_cases("b" * 24)[0]["status"], "fail")
        self.save_run("b" * 24, result={**self.result, "rows": [self.row, self.row]})
        self.assertEqual(self.lab.evaluate_cases("b" * 24)[0]["status"], "fail")

    def test_retention_null_is_distinct_from_zero_and_bad_expected_values_rejected(self):
        self.case(metric="ttm_gross_retention", expected=None)
        self.assertEqual(self.lab.evaluate_cases(self.identity)[0]["status"], "pass")
        self.save_run("b" * 24, result={**self.result, "rows": [{**self.row, "ttm_gross_retention": 0}]})
        self.assertEqual(self.lab.evaluate_cases("b" * 24)[0]["status"], "fail")
        for expected in (True, float("nan"), -1, None):
            with self.subTest(expected=expected), self.assertRaises(ValueError):
                self.case(expected=expected)
        with self.assertRaises(ValueError):
            self.case(metric="gross_retention", expected=85)

    def test_regression_rerun_keeps_answer_hidden_and_verifies_fingerprint(self):
        case = self.case()
        captured = []
        def create(data, name, **kwargs):
            captured.append((data, name, kwargs))
            return {"id": "new"}
        self.lab.rerun_case(case["id"], {"model": "test-b"}, create)
        self.assertEqual(captured[0][0], self.data)
        self.assertNotIn("expected", captured[0][2])
        self.assertNotIn("rationale", captured[0][2])
        self.assertEqual(captured[0][2]["selected_model"], "test-b")
        path = self.root / "cases" / case["id"] / "case.json"
        changed = self.lab.read(path)
        changed["expected"] = 999
        self.lab.write(path, changed)
        with self.assertRaisesRegex(ValueError, "fingerprint"):
            self.lab.rerun_case(case["id"], {"model": "test-b"}, create)

    def test_configuration_is_frozen_before_worker_and_no_inline_key_is_saved(self):
        with patch.object(app, "STORE", self.root), patch.object(app.POOL, "submit"), patch.object(app, "load_config", return_value={
            "provider": "openai_responses", "model": "default", "api_key": "private", "max_output_tokens": 500}):
            run = app.create_run(self.data, "example.xlsx", mode="model", selected_model="test-b")
            meta = self.lab.read(self.root / run["id"] / "run.json")
        self.assertEqual(meta["model_config"]["model"], "test-b")
        self.assertNotIn("api_key", meta["model_config"])
        self.assertEqual(meta["config_sha256"], fingerprint(meta["model_config"]))

    def test_row_absence_regression_fails_original_extra_row_and_passes_corrected_output(self):
        case = self.case(assertion="row_absent", metric=None, expected=None)
        result = self.lab.evaluate_cases(self.identity)[0]
        self.assertEqual(result["status"], "fail")
        self.assertEqual(result["actual"], 1)
        self.save_run("b" * 24, result={**self.result, "rows": []})
        result = self.lab.evaluate_cases("b" * 24)[0]
        self.assertEqual(result["status"], "pass")
        self.assertEqual(result["actual"], 0)
        self.assertIsNone(case["metric"])

    def test_row_presence_assertion_rejects_duplicate_rows(self):
        self.case(assertion="row_present", metric=None, expected=None)
        self.assertEqual(self.lab.evaluate_cases(self.identity)[0]["status"], "pass")
        self.save_run("b" * 24, result={**self.result, "rows": [self.row, self.row]})
        self.assertEqual(self.lab.evaluate_cases("b" * 24)[0]["status"], "fail")


if __name__ == "__main__":
    unittest.main()

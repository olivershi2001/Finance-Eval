"""Behavioral checks for the upload -> analyze -> review -> export workflow."""
from dataclasses import replace
from datetime import date
import hashlib
from http.server import ThreadingHTTPServer
import io
import json
from pathlib import Path
import tempfile
import threading
import unittest
from unittest.mock import patch
from urllib.error import HTTPError
from urllib.request import Request, urlopen
import zipfile

import app
from models import CubeConfig, RevenueRecord
from customer_cube import calculate_cube
from xlsx_reader import load_revenue_records


class WorkbenchTests(unittest.TestCase):
    def setUp(self):
        self.temp = tempfile.TemporaryDirectory()
        self.store_patch = patch.object(app, "STORE", Path(self.temp.name))
        self.store_patch.start()
        self.data = (app.ROOT / "fixtures/customer_cube_raw.xlsx").read_bytes()
        self.submit = patch.object(app.POOL, "submit")
        self.submit.start()

    def tearDown(self):
        self.submit.stop()
        self.store_patch.stop()
        self.temp.cleanup()

    def create(self, **kwargs):
        run = app.create_run(self.data, "example.xlsx", **kwargs)
        return run, app.run_dir(run["id"])

    def test_calculated_analysis_preserves_source_and_complete_reference_rows(self):
        run, path = self.create()
        with patch.object(app, "export_workbook"):
            app.process_run(run["id"])
        meta = app.read_json(path / "run.json")
        result = app.read_json(path / "result.json")
        self.assertEqual(meta["status"], "complete")
        self.assertEqual(result["source"]["row_count"], 1000)
        self.assertEqual(len(result["rows"]), 604)
        self.assertIsNone(result["score"])
        self.assertEqual(meta["source_sha256"], hashlib.sha256(self.data).hexdigest())
        self.assertEqual(app.read_json(path / "source_evidence.json")[0]["source_row"], 2)
        self.assertTrue((path / "analysis.csv").is_file())

    def test_failed_ai_explanation_keeps_calculations_and_exposes_warning(self):
        run, path = self.create(mode="assisted")
        with patch.object(app, "generate_narrative", side_effect=ValueError("No API key configured")), patch.object(app, "export_workbook"):
            app.process_run(run["id"])
        result = app.read_json(path / "result.json")
        self.assertEqual(app.read_json(path / "run.json")["status"], "complete")
        self.assertEqual(len(result["rows"]), 604)
        self.assertIn("No API key", result["warnings"][0])
        self.assertIn("no AI call", result["narrative"]["origin"])

    def test_export_failure_does_not_lose_analysis(self):
        run, path = self.create()
        with patch.object(app, "export_workbook", side_effect=ValueError("No workbook runtime")):
            app.process_run(run["id"])
        self.assertEqual(app.read_json(path / "run.json")["status"], "complete")
        self.assertIn("Excel export unavailable", app.read_json(path / "result.json")["warnings"][0])
        self.assertFalse(app.read_json(path / "run.json")["has_excel"])

    def test_model_path_scores_real_candidate_errors(self):
        run, path = self.create(mode="model")
        candidate = json.loads((app.ROOT / "tests/fixtures/candidate_model_response.json").read_text())
        with patch.object(app, "call_candidate", return_value=json.dumps(candidate)), patch.object(app, "export_workbook"):
            app.process_run(run["id"])
        self.assertEqual(app.read_json(path / "run.json")["status"], "complete")
        result = app.read_json(path / "result.json")
        self.assertIsNotNone(result["score"])
        self.assertLess(result["score"]["score"], 100)
        self.assertTrue((path / "candidate_response.txt").is_file())

    def test_narrative_rejects_nonexistent_evidence_citation(self):
        run, path = self.create()
        rows = calculate_cube(load_revenue_records(path / "source.xlsx"), CubeConfig())
        response = {"summary": "Test", "insights": [{"title": "Invalid citation", "detail": "test", "row_keys": [["quarterly", "2099-Q1", "overall", "All"]]}], "questions": []}
        with patch.object(app, "call_openai_responses", return_value=json.dumps(response)):
            with self.assertRaisesRegex(ValueError, "does not exist"):
                app.generate_narrative(path, rows, [], {"provider": "openai_responses", "model": "test", "max_output_tokens": 500}, "period_sum")

    def test_rate_comparison_does_not_allow_one_percentage_point_errors(self):
        records = load_revenue_records(app.ROOT / "fixtures/customer_cube_raw.xlsx")
        row = next(r for r in calculate_cube(records, CubeConfig()) if r.net_retention is not None)
        changed = replace(row, net_retention=row.net_retention + .005)
        differences = app.compare_rows([row], [changed], CubeConfig(absolute_tolerance=1e-6))
        self.assertEqual(differences[0]["metric"], "net_retention")

    def test_invalid_amounts_fail_instead_of_returning_a_plausible_analysis(self):
        for amount in (float("nan"), float("inf"), -1):
            with self.subTest(amount=amount):
                run, path = self.create()
                records = (RevenueRecord("Example", date(2025, 1, 1), amount, "Platform"),)
                with patch.object(app, "load_revenue_records", return_value=records):
                    app.process_run(run["id"])
                self.assertEqual(app.read_json(path / "run.json")["status"], "failed")
                self.assertFalse((path / "result.json").exists())

    def test_model_failure_is_not_reported_as_a_successful_calculated_run(self):
        run, path = self.create(mode="model")
        with patch.object(app, "call_candidate", side_effect=ValueError("Model refused")):
            app.process_run(run["id"])
        self.assertEqual(app.read_json(path / "run.json")["status"], "failed")
        self.assertFalse((path / "result.json").exists())

    def test_judge_receives_real_workbook_path_and_retains_validated_result(self):
        run, path = self.create(judge=True)
        def export(*args):
            (path / "analysis.xlsx").write_bytes(b"mock export")
        with patch.object(app, "export_workbook", side_effect=export), patch.object(app, "build_judge_prompt", return_value="judge input") as prompt, patch.object(app, "call_judge", return_value='{"test":true}'), patch.object(app, "validate_judge_response", return_value={"validated_score": 87, "outcome": "PASS"}):
            app.process_run(run["id"])
        self.assertEqual(prompt.call_args.args[1], path / "analysis.xlsx")
        self.assertEqual(app.read_json(path / "result.json")["judge"]["validated_score"], 87)
        self.assertEqual(app.read_json(path / "run.json")["status"], "complete")

    def test_source_evidence_keeps_sparse_excel_row_numbers(self):
        from xlsx_reader import WorkbookCell
        run, path = self.create()
        records = (RevenueRecord("Example", date(2025, 1, 1), 10, "Platform"),)
        cells = {"A1": WorkbookCell("Customer name"), "A2": WorkbookCell(""),
                 "A8": WorkbookCell("Example"), "B8": WorkbookCell("2025-01"),
                 "C8": WorkbookCell(10), "D8": WorkbookCell("Platform")}
        with patch.object(app, "read_workbook_cells", return_value={"Original": cells}):
            evidence = app.source_evidence(path / "source.xlsx", records)
        self.assertEqual(evidence[0]["source_row"], 8)
        self.assertEqual(evidence[0]["sheet"], "Original")

    def test_rejects_invalid_uploads_and_incompatible_conventions(self):
        with self.assertRaisesRegex(ValueError, "xlsx"):
            app.create_run(b"not an excel file", "bad.xlsx")
        with self.assertRaisesRegex(ValueError, "monthly revenue"):
            self.create(mode="model", amount_mode="period_end")
        buffer = io.BytesIO()
        with zipfile.ZipFile(buffer, "w") as archive:
            archive.writestr("xl/workbook.xml", "")
            archive.writestr("oversize", b"x" * (31 * 1024 * 1024), compress_type=zipfile.ZIP_DEFLATED)
        with self.assertRaisesRegex(ValueError, "size"):
            app.validate_upload(buffer.getvalue())

    def test_csv_cells_and_provider_errors_do_not_expose_executable_text_or_keys(self):
        self.assertEqual(app.csv_safe("=HYPERLINK(\"test\")"), "'=HYPERLINK(\"test\")")
        self.assertEqual(app.csv_safe(-50), -50)
        self.assertNotIn("sk-secret", app.safe_error(ValueError("OpenAI API returned HTTP 401: sk-secret")))

    def test_http_upload_download_feedback_and_origin_protection(self):
        server = ThreadingHTTPServer(("127.0.0.1", 0), app.Handler)
        thread = threading.Thread(target=server.serve_forever, daemon=True)
        thread.start()
        base = f"http://127.0.0.1:{server.server_port}"
        try:
            req = Request(base + "/api/runs?demo=1", data=b"", method="POST")
            with urlopen(req) as response:
                self.assertEqual(response.status, 202)
                run = json.load(response)
            with patch.object(app, "export_workbook"):
                app.process_run(run["id"])
            url = base + "/api/runs/" + run["id"]
            with urlopen(url) as response:
                result = json.load(response)["result"]
            row_key = [result["rows"][0][k] for k in ("grain", "period", "slice_type", "slice_value")]
            req = Request(url + "/feedback", data=json.dumps({"note": "Check the first period", "status": "flagged", "row_key": row_key}).encode(), headers={"Content-Type": "application/json"})
            with urlopen(req) as response:
                self.assertEqual(json.load(response)[0]["row_key"], row_key)
            with urlopen(url + "/bundle") as response:
                with zipfile.ZipFile(io.BytesIO(response.read())) as archive:
                    self.assertIn("feedback.json", archive.namelist())
                    self.assertIn("analysis.csv", archive.namelist())
                    self.assertEqual(json.loads(archive.read("feedback.json"))[0]["note"], "Check the first period")
            with self.assertRaises(HTTPError) as error:
                urlopen(Request(base + "/api/runs?demo=1", data=b"", headers={"Origin": "https://other.example"}))
            self.assertEqual(error.exception.code, 403)
            error.exception.close()
            with self.assertRaises(HTTPError) as error:
                urlopen(Request(url + "/feedback", data=json.dumps({"note": "test", "row_key": ["bad"]}).encode()))
            self.assertEqual(error.exception.code, 400)
            error.exception.close()
        finally:
            server.shutdown()
            server.server_close()
            thread.join(timeout=2)


if __name__ == "__main__":
    unittest.main()

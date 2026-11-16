"""Tests for the answer-key-free LLM workbook judge."""

from __future__ import annotations

import json
import sys
import unittest
from pathlib import Path
from unittest.mock import Mock, patch


ROOT = Path(__file__).resolve().parents[1]
sys.path.insert(0, str(ROOT))

from llm_judge import (  # noqa: E402
    _extract_openai_output_text,
    build_judge_prompt,
    call_openai_responses,
    load_config,
    load_rubric,
    resolve_openai_api_key,
    validate_judge_response,
    workbook_to_evidence,
)


class LlmJudgeTests(unittest.TestCase):
    @classmethod
    def setUpClass(cls) -> None:
        cls.rubric = load_rubric(
            ROOT / "rubrics" / "customer_cube_llm_judge.json"
        )
        cls.config = load_config(ROOT / "llm_judge_config.json")

    def response_with_scores(self, deduction: float = 0.0) -> dict[str, object]:
        criteria = []
        for index, criterion in enumerate(self.rubric["criteria"]):
            score = float(criterion["max_points"])
            if index == 0:
                score -= deduction
            criteria.append(
                {
                    "id": criterion["id"],
                    "score": score,
                    "rationale": "The sampled values and reported totals support this score.",
                    "evidence": ["Quarterly!A1:M20 - sampled and reconciled output"],
                }
            )
        return {
            "summary": "The workbook is complete and materially correct.",
            "confidence": 0.9,
            "critical_errors": [],
            "criteria": criteria,
        }

    def test_rubric_is_a_complete_one_hundred_point_scorecard(self) -> None:
        maximum = sum(item["max_points"] for item in self.rubric["criteria"])

        self.assertEqual(maximum, 100)
        self.assertEqual(len(self.rubric["criteria"]), 6)

    def test_imported_rubric_overrides_weights_and_keeps_definitions(self) -> None:
        rubric = load_rubric(
            ROOT / "rubrics" / "customer_cube_strict_accuracy.json"
        )
        weights = {item["id"]: item["max_points"] for item in rubric["criteria"]}

        self.assertEqual(rubric["rubric_id"], "customer_cube_strict_accuracy")
        self.assertEqual(rubric["pass_threshold"], 90)
        self.assertEqual(weights["movement_classification"], 40)
        self.assertEqual(weights["retention"], 25)
        self.assertEqual(sum(weights.values()), 100)
        self.assertIn("new", rubric["metric_definitions"])

    def test_all_bundled_rubrics_resolve_to_one_hundred_points(self) -> None:
        paths = sorted((ROOT / "rubrics").glob("*.json"))

        self.assertGreaterEqual(len(paths), 4)
        for path in paths:
            with self.subTest(path=path.name):
                rubric = load_rubric(path)
                maximum = sum(item["max_points"] for item in rubric["criteria"])
                self.assertEqual(maximum, 100)

    def test_finance_template_rubric_preserves_anchored_scale_and_thresholds(
        self,
    ) -> None:
        rubric = load_rubric(
            ROOT / "rubrics" / "customer_cube_finance_template.json"
        )

        self.assertEqual(rubric["rating_scale"]["minimum"], 0)
        self.assertEqual(rubric["rating_scale"]["maximum"], 4)
        self.assertEqual(rubric["pass_threshold"], 85)
        self.assertEqual(rubric["review_threshold"], 70)
        self.assertEqual(len(rubric["criteria"]), 7)
        self.assertEqual(sum(item["max_points"] for item in rubric["criteria"]), 100)

    def test_anchored_ratings_are_converted_to_weighted_points(self) -> None:
        rubric = load_rubric(
            ROOT / "rubrics" / "customer_cube_finance_template.json"
        )
        ratings = {"C1": 4, "C2": 4, "C3": 4, "C4": 3, "C5": 1, "C6": 3, "C7": 3}
        response = {
            "summary": "The workbook is useful but needs normal senior review.",
            "confidence": 0.85,
            "critical_errors": [],
            "criteria": [
                {
                    "id": criterion["id"],
                    "rating": ratings[criterion["id"]],
                    "rationale": "The workbook evidence supports this anchored rating.",
                    "evidence": ["Quarterly!A1:N20 - sampled evidence"],
                }
                for criterion in rubric["criteria"]
            ],
        }

        report = validate_judge_response(response, rubric)

        self.assertEqual(report["validated_score"], 81.25)
        self.assertEqual(report["outcome"], "REVIEW")
        self.assertFalse(report["passed"])
        self.assertEqual(report["criteria"][0]["rating"], 4)
        self.assertEqual(report["criteria"][0]["score"], 30)

    def test_anchored_rating_rejects_values_between_declared_steps(self) -> None:
        rubric = load_rubric(
            ROOT / "rubrics" / "customer_cube_finance_template.json"
        )
        response = {
            "summary": "Invalid half-step response.",
            "confidence": 0.8,
            "critical_errors": [],
            "criteria": [
                {
                    "id": criterion["id"],
                    "rating": 2.5 if index == 0 else 4,
                    "rationale": "Evidence was reviewed.",
                    "evidence": ["Annual!A1:N20 - sampled evidence"],
                }
                for index, criterion in enumerate(rubric["criteria"])
            ],
        }

        with self.assertRaisesRegex(ValueError, "increments of 1"):
            validate_judge_response(response, rubric)

    def test_raw_workbook_is_serialized_with_addresses_and_all_rows(self) -> None:
        evidence = workbook_to_evidence(
            ROOT / "fixtures" / "customer_cube_raw.xlsx",
            include_formulas=True,
            max_evidence_cells=10_000,
        )

        self.assertEqual(evidence["cell_count"], 4_004)
        self.assertEqual(evidence["sheets"][0]["used_range"], "A1:D1001")
        self.assertEqual(len(evidence["sheets"][0]["rows"]), 1_001)

    def test_workbook_evidence_includes_observable_number_formats(self) -> None:
        evidence = workbook_to_evidence(
            ROOT / "fixtures" / "customer_cube_raw.xlsx",
            include_formulas=True,
            max_evidence_cells=10_000,
        )
        first_data_row = evidence["sheets"][0]["rows"][1]["values"]

        self.assertEqual(first_data_row[1]["number_format"], "mmm yyyy")
        self.assertIn("$", first_data_row[2]["number_format"])

    def test_prompt_contains_rubric_security_rule_and_both_workbooks(self) -> None:
        prompt = build_judge_prompt(
            ROOT / "fixtures" / "customer_cube_raw.xlsx",
            ROOT
            / "outputs"
            / "customer_cube_eval"
            / "customer_cube_reference_output.xlsx",
            self.rubric,
            self.config,
        )

        self.assertIn("customer_cube_llm_judge", prompt)
        self.assertIn("untrusted", prompt.lower())
        self.assertIn("customer_cube_raw.xlsx", prompt)
        self.assertIn("customer_cube_reference_output.xlsx", prompt)
        self.assertIn("Do not calculate or return an overall score", prompt)

    def test_validator_calculates_score_and_pass_instead_of_trusting_model(self) -> None:
        report = validate_judge_response(
            self.response_with_scores(deduction=5), self.rubric
        )

        self.assertEqual(report["validated_score"], 95.0)
        self.assertTrue(report["passed"])

    def test_critical_error_forces_failure_even_above_threshold(self) -> None:
        response = self.response_with_scores()
        response["critical_errors"] = [
            {
                "id": "pervasive_reconciliation_failure",
                "evidence": "Quarterly!A1:M200 - the bridge fails throughout.",
            }
        ]

        report = validate_judge_response(response, self.rubric)

        self.assertEqual(report["validated_score"], 100.0)
        self.assertFalse(report["passed"])

    def test_validator_rejects_an_invented_critical_failure(self) -> None:
        response = self.response_with_scores()
        response["critical_errors"] = [
            {
                "id": "ugly_formatting",
                "evidence": "Summary!A1 - the colors are unattractive.",
            }
        ]

        with self.assertRaisesRegex(ValueError, "unknown critical failure"):
            validate_judge_response(response, self.rubric)

    def test_validator_rejects_missing_criterion(self) -> None:
        response = self.response_with_scores()
        response["criteria"].pop()

        with self.assertRaisesRegex(ValueError, "missing criteria"):
            validate_judge_response(response, self.rubric)

    def test_openai_response_text_is_extracted_from_response_items(self) -> None:
        response = {
            "output": [
                {
                    "type": "message",
                    "content": [
                        {"type": "output_text", "text": '{"summary":"ok"}'}
                    ],
                }
            ]
        }

        self.assertEqual(_extract_openai_output_text(response), '{"summary":"ok"}')

    def test_api_key_can_be_loaded_from_macos_keychain(self) -> None:
        config = {
            "api_key_env": "TEST_OPENAI_API_KEY",
            "keychain_service": "com.example.test-key",
        }
        completed = Mock(returncode=0, stdout="sk-test-key\n")

        with (
            patch.dict("os.environ", {}, clear=True),
            patch("llm_judge.sys.platform", "darwin"),
            patch("llm_judge.subprocess.run", return_value=completed) as run,
        ):
            key = resolve_openai_api_key(config)

        self.assertEqual(key, "sk-test-key")
        run.assert_called_once_with(
            [
                "/usr/bin/security",
                "find-generic-password",
                "-s",
                "com.example.test-key",
                "-w",
            ],
            text=True,
            capture_output=True,
            check=False,
        )

    def test_openai_request_can_require_a_strict_json_schema(self) -> None:
        config = {
            "model": "test-model",
            "api_url": "https://api.openai.com/v1/responses",
            "api_key_env": "OPENAI_API_KEY",
            "keychain_service": "",
            "max_output_tokens": 1000,
            "timeout_seconds": 30,
            "store": False,
        }
        schema = {
            "type": "object",
            "properties": {"answer": {"type": "string"}},
            "required": ["answer"],
            "additionalProperties": False,
        }
        api_response = {
            "output": [
                {
                    "type": "message",
                    "content": [
                        {"type": "output_text", "text": '{"answer":"ok"}'}
                    ],
                }
            ]
        }
        opened = Mock()
        opened.__enter__ = Mock(return_value=opened)
        opened.__exit__ = Mock(return_value=False)
        opened.read.return_value = json.dumps(api_response).encode("utf-8")

        with (
            patch("llm_judge.resolve_openai_api_key", return_value="sk-test"),
            patch("llm_judge.urllib.request.urlopen", return_value=opened) as urlopen,
        ):
            response = call_openai_responses(
                "Return JSON.",
                config,
                response_schema=schema,
                schema_name="answer_schema",
            )

        self.assertEqual(response, '{"answer":"ok"}')
        request = urlopen.call_args.args[0]
        body = json.loads(request.data.decode("utf-8"))
        self.assertFalse(body["store"])
        self.assertEqual(body["text"]["format"]["type"], "json_schema")
        self.assertEqual(body["text"]["format"]["name"], "answer_schema")
        self.assertTrue(body["text"]["format"]["strict"])
        self.assertEqual(body["text"]["format"]["schema"], schema)


if __name__ == "__main__":
    unittest.main()

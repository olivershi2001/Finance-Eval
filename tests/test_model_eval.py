"""Tests for the end-to-end model-under-test workflow."""

from __future__ import annotations

import json
import sys
import unittest
from pathlib import Path


ROOT = Path(__file__).resolve().parents[1]
sys.path.insert(0, str(ROOT))

from llm_judge import load_config  # noqa: E402
from model_eval import (  # noqa: E402
    build_candidate_prompt,
    candidate_response_schema,
    validate_candidate_response,
)


class ModelEvalTests(unittest.TestCase):
    @classmethod
    def setUpClass(cls) -> None:
        cls.response = json.loads(
            (ROOT / "tests" / "fixtures" / "candidate_model_response.json").read_text(
                encoding="utf-8"
            )
        )

    def test_schema_is_strict_and_contains_both_period_grains(self) -> None:
        schema = candidate_response_schema()

        self.assertFalse(schema["additionalProperties"])
        self.assertIn("quarterly_rows", schema["required"])
        self.assertIn("annual_rows", schema["required"])
        row_schema = schema["properties"]["quarterly_rows"]["items"]
        self.assertFalse(row_schema["additionalProperties"])
        self.assertEqual(row_schema["properties"]["slice_type"]["enum"], [
            "overall",
            "product",
            "customer",
        ])

    def test_saved_model_response_validates_without_a_candidate_workbook(self) -> None:
        validated = validate_candidate_response(self.response)

        self.assertEqual(len(validated["quarterly_rows"]), 2)
        self.assertEqual(len(validated["annual_rows"]), 1)
        self.assertEqual(validated["quarterly_rows"][0]["net_retention"], 1.3)

    def test_validator_rejects_missing_required_output_section(self) -> None:
        response = dict(self.response)
        response.pop("annual_rows")

        with self.assertRaisesRegex(ValueError, "annual_rows"):
            validate_candidate_response(response)

    def test_prompt_treats_workbook_cells_as_untrusted_data(self) -> None:
        config = load_config(ROOT / "candidate_model_config.json")
        prompt = build_candidate_prompt(
            ROOT / "fixtures" / "customer_cube_raw.xlsx", config
        )

        self.assertIn("model being evaluated", prompt)
        self.assertIn("untrusted", prompt.lower())
        self.assertIn("quarterly_rows", prompt)
        self.assertNotIn("pass_threshold", prompt)


if __name__ == "__main__":
    unittest.main()

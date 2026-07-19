"""CLI and reusable runner for the customer-cube evaluation."""

from __future__ import annotations

import argparse
import json
from dataclasses import replace
from pathlib import Path
from typing import Any

from candidate import build_customer_cube
from models import CandidateResult, CubeConfig, CubeRow, EvalReport, SliceType
from scoring import score_cube
from xlsx_reader import load_revenue_records


ROOT = Path(__file__).parent
DEFAULT_INPUT = ROOT / "fixtures" / "customer_cube_raw.xlsx"
DEFAULT_EXPECTED = ROOT / "scenarios" / "customer_cube_expected.json"
DEFAULT_CONFIG = ROOT / "eval_config.json"


def load_config(path: Path = DEFAULT_CONFIG) -> CubeConfig:
    with path.open(encoding="utf-8") as file:
        data = json.load(file)
    return CubeConfig(
        amount_mode=str(data.get("amount_mode", "period_sum")),  # type: ignore[arg-type]
        fiscal_year_start_month=int(data.get("fiscal_year_start_month", 1)),
        slices=tuple(data.get("slices", ["overall", "product", "customer"])),
        absolute_tolerance=float(data.get("absolute_tolerance", 0.01)),
        relative_tolerance=float(data.get("relative_tolerance", 1e-6)),
        pass_threshold=float(data.get("pass_threshold", 95.0)),
    )


def load_cube_rows(path: Path) -> tuple[CubeRow, ...]:
    with path.open(encoding="utf-8") as file:
        payload = json.load(file)
    raw_rows = payload["rows"] if isinstance(payload, dict) else payload
    if not isinstance(raw_rows, list):
        raise ValueError("Cube JSON must be a list or an object with a rows list")
    return tuple(CubeRow.from_dict(row) for row in raw_rows)


def save_cube_rows(path: Path, rows: tuple[CubeRow, ...]) -> None:
    path.parent.mkdir(parents=True, exist_ok=True)
    payload = {
        "contract_version": 1,
        "rows": [row.to_dict() for row in rows],
    }
    path.write_text(json.dumps(payload, indent=2) + "\n", encoding="utf-8")


def run_eval(
    input_path: Path = DEFAULT_INPUT,
    expected_path: Path = DEFAULT_EXPECTED,
    config_path: Path = DEFAULT_CONFIG,
) -> EvalReport:
    config = load_config(config_path)
    records = load_revenue_records(input_path)
    candidate_result = build_customer_cube(records, config)
    if not isinstance(candidate_result, CandidateResult):
        raise TypeError("Candidate must return CandidateResult")
    return score_cube(load_cube_rows(expected_path), candidate_result.rows, config)


def _print_report(report: EvalReport) -> None:
    status = "PASS" if report.passed else "FAIL"
    print(f"[{status}] Customer cube score: {report.score:.2f}/100")
    print(
        "  structure={:.2f}/10 coverage={:.2f}/20 movements={:.2f}/40 "
        "retention={:.2f}/20 reconciliation={:.2f}/10".format(
            report.structure_score,
            report.coverage_score,
            report.movement_score,
            report.retention_score,
            report.reconciliation_score,
        )
    )
    print(f"  rows: expected={report.expected_rows} actual={report.actual_rows}")
    for slice_score in report.slice_scores:
        print(
            f"  {slice_score.grain}/{slice_score.slice_type}: "
            f"{100 * slice_score.accuracy:.1f}% cell accuracy"
        )
    if report.missing_keys:
        print(f"  missing keys (first {len(report.missing_keys)}):")
        for key in report.missing_keys:
            print(f"    {key}")
    if report.extra_keys:
        print(f"  extra keys (first {len(report.extra_keys)}):")
        for key in report.extra_keys:
            print(f"    {key}")
    if report.duplicate_keys:
        print(f"  duplicate keys (first {len(report.duplicate_keys)}):")
        for key in report.duplicate_keys:
            print(f"    {key}")
    if report.mismatches:
        print(f"  value mismatches (first {len(report.mismatches)}):")
        for mismatch in report.mismatches:
            print(f"    {mismatch}")


def _add_common_paths(parser: argparse.ArgumentParser) -> None:
    parser.add_argument("--input", type=Path, default=DEFAULT_INPUT)
    parser.add_argument("--config", type=Path, default=DEFAULT_CONFIG)


def _build_parser() -> argparse.ArgumentParser:
    parser = argparse.ArgumentParser(description=__doc__)
    subparsers = parser.add_subparsers(dest="command")

    evaluate_parser = subparsers.add_parser("evaluate", help="Grade candidate.py")
    _add_common_paths(evaluate_parser)
    evaluate_parser.add_argument("--expected", type=Path, default=DEFAULT_EXPECTED)

    analyze_parser = subparsers.add_parser(
        "analyze", help="Run the reference candidate on any four-column workbook"
    )
    _add_common_paths(analyze_parser)
    analyze_parser.add_argument("--output", type=Path, required=True)
    analyze_parser.add_argument(
        "--slice",
        action="append",
        choices=("overall", "product", "customer"),
        dest="slices",
        help="Include this cut; repeat for multiple cuts (default: all)",
    )

    score_parser = subparsers.add_parser(
        "score", help="Grade a saved candidate JSON without editing candidate.py"
    )
    score_parser.add_argument("candidate_output", type=Path)
    score_parser.add_argument("--expected", type=Path, default=DEFAULT_EXPECTED)
    score_parser.add_argument("--config", type=Path, default=DEFAULT_CONFIG)
    return parser


def main() -> int:
    parser = _build_parser()
    args = parser.parse_args()
    command = args.command or "evaluate"

    if command == "evaluate":
        report = run_eval(
            getattr(args, "input", DEFAULT_INPUT),
            getattr(args, "expected", DEFAULT_EXPECTED),
            getattr(args, "config", DEFAULT_CONFIG),
        )
        _print_report(report)
        return 0 if report.passed else 1

    if command == "analyze":
        config = load_config(args.config)
        if args.slices:
            config = replace(
                config,
                slices=tuple(args.slices),  # type: ignore[arg-type]
            )
        rows = build_customer_cube(
            load_revenue_records(args.input), config
        ).rows
        save_cube_rows(args.output, rows)
        print(f"Wrote {len(rows)} rows to {args.output}")
        return 0

    if command == "score":
        config = load_config(args.config)
        report = score_cube(
            load_cube_rows(args.expected),
            load_cube_rows(args.candidate_output),
            config,
        )
        _print_report(report)
        return 0 if report.passed else 1

    parser.error(f"Unknown command: {command}")
    return 2


if __name__ == "__main__":
    raise SystemExit(main())

"""Generic, configurable XLSX-against-XLSX evaluation CLI.

The expected workbook is the answer key. By default, every nonblank expected
cell is scored against the same sheet/cell coordinate in the candidate workbook.
Formatting is deliberately ignored.
"""

from __future__ import annotations

import argparse
import json
import math
import re
from dataclasses import asdict, dataclass, field
from datetime import date, datetime
from pathlib import Path
from typing import Literal

from xlsx_reader import WorkbookCell, WorkbookCells, read_workbook_cells


ROOT = Path(__file__).parent
DEFAULT_CONFIG = ROOT / "workbook_score_config.json"
FormulaMode = Literal["values", "exact", "values_and_formulas"]


@dataclass(frozen=True, slots=True)
class CellRange:
    start_row: int
    start_col: int
    end_row: int
    end_col: int

    def contains(self, row: int, col: int) -> bool:
        return (
            self.start_row <= row <= self.end_row
            and self.start_col <= col <= self.end_col
        )


@dataclass(frozen=True, slots=True)
class WeightedRange:
    sheet: str
    cell_range: CellRange
    weight: float


@dataclass(frozen=True, slots=True)
class WorkbookScoreConfig:
    include_sheets: tuple[str, ...] = ()
    exclude_sheets: tuple[str, ...] = ()
    sheet_map: dict[str, str] = field(default_factory=dict)
    ranges: dict[str, tuple[CellRange, ...]] = field(default_factory=dict)
    ignore_ranges: dict[str, tuple[CellRange, ...]] = field(default_factory=dict)
    sheet_weights: dict[str, float] = field(default_factory=dict)
    range_weights: tuple[WeightedRange, ...] = ()
    absolute_tolerance: float = 0.01
    relative_tolerance: float = 1e-6
    text_case_sensitive: bool = False
    collapse_whitespace: bool = True
    sheet_name_case_sensitive: bool = False
    formula_mode: FormulaMode = "values"
    score_blank_expected_cells: bool = False
    penalize_extra_cells: bool = False
    penalize_extra_sheets: bool = False
    pass_threshold: float = 95.0
    max_mismatches: int = 50


@dataclass(frozen=True, slots=True)
class SheetScore:
    expected_sheet: str
    candidate_sheet: str | None
    expected_cells: int
    covered_cells: int
    matched_cells: int
    extra_cells: int
    weighted_accuracy: float


@dataclass(frozen=True, slots=True)
class WorkbookScoreReport:
    score: float
    passed: bool
    structure_score: float
    coverage_score: float
    value_score: float
    expected_sheets: int
    matched_sheets: int
    expected_cells: int
    covered_cells: int
    matched_cells: int
    extra_cells: int
    extra_sheets: tuple[str, ...]
    sheet_scores: tuple[SheetScore, ...]
    mismatches: tuple[str, ...]

    def to_dict(self) -> dict[str, object]:
        return asdict(self)


CELL_REFERENCE = re.compile(r"^\$?([A-Z]+)\$?(\d+)$", re.IGNORECASE)
RANGE_REFERENCE = re.compile(
    r"^\s*\$?([A-Z]+)\$?(\d+)\s*:\s*\$?([A-Z]+)\$?(\d+)\s*$",
    re.IGNORECASE,
)


def _column_number(letters: str) -> int:
    result = 0
    for character in letters.upper():
        result = result * 26 + ord(character) - ord("A") + 1
    return result


def _column_letters(number: int) -> str:
    result = ""
    while number:
        number, remainder = divmod(number - 1, 26)
        result = chr(ord("A") + remainder) + result
    return result


def _cell_coordinates(reference: str) -> tuple[int, int]:
    match = CELL_REFERENCE.fullmatch(reference)
    if not match:
        raise ValueError(f"Invalid cell reference: {reference!r}")
    return int(match.group(2)), _column_number(match.group(1))


def _parse_range(reference: str) -> CellRange:
    cell = CELL_REFERENCE.fullmatch(reference.strip())
    if cell:
        row = int(cell.group(2))
        col = _column_number(cell.group(1))
        return CellRange(row, col, row, col)
    match = RANGE_REFERENCE.fullmatch(reference)
    if not match:
        raise ValueError(f"Invalid A1 range: {reference!r}")
    start_col = _column_number(match.group(1))
    start_row = int(match.group(2))
    end_col = _column_number(match.group(3))
    end_row = int(match.group(4))
    return CellRange(
        min(start_row, end_row),
        min(start_col, end_col),
        max(start_row, end_row),
        max(start_col, end_col),
    )


def _parse_range_map(data: object, name: str) -> dict[str, tuple[CellRange, ...]]:
    if data is None:
        return {}
    if not isinstance(data, dict):
        raise ValueError(f"{name} must be an object keyed by expected sheet name")
    result: dict[str, tuple[CellRange, ...]] = {}
    for sheet, references in data.items():
        if not isinstance(references, list):
            raise ValueError(f"{name}.{sheet} must be a list of A1 ranges")
        result[str(sheet)] = tuple(_parse_range(str(item)) for item in references)
    return result


def load_workbook_score_config(path: Path = DEFAULT_CONFIG) -> WorkbookScoreConfig:
    with path.open(encoding="utf-8") as file:
        data = json.load(file)
    formula_mode = str(data.get("formula_mode", "values"))
    if formula_mode not in ("values", "exact", "values_and_formulas"):
        raise ValueError(
            "formula_mode must be values, exact, or values_and_formulas"
        )
    weighted_ranges: list[WeightedRange] = []
    for item in data.get("range_weights", []):
        if not isinstance(item, dict):
            raise ValueError("range_weights entries must be objects")
        weight = float(item["weight"])
        if weight <= 0:
            raise ValueError("range weight must be positive")
        weighted_ranges.append(
            WeightedRange(
                sheet=str(item["sheet"]),
                cell_range=_parse_range(str(item["range"])),
                weight=weight,
            )
        )
    sheet_weights = {
        str(sheet): float(weight)
        for sheet, weight in data.get("sheet_weights", {}).items()
    }
    if any(weight <= 0 for weight in sheet_weights.values()):
        raise ValueError("sheet weights must be positive")
    absolute_tolerance = float(data.get("absolute_tolerance", 0.01))
    relative_tolerance = float(data.get("relative_tolerance", 1e-6))
    pass_threshold = float(data.get("pass_threshold", 95.0))
    max_mismatches = int(data.get("max_mismatches", 50))
    if absolute_tolerance < 0 or relative_tolerance < 0:
        raise ValueError("numeric tolerances cannot be negative")
    if not 0 <= pass_threshold <= 100:
        raise ValueError("pass_threshold must be between 0 and 100")
    if max_mismatches < 0:
        raise ValueError("max_mismatches cannot be negative")
    return WorkbookScoreConfig(
        include_sheets=tuple(str(item) for item in data.get("include_sheets", [])),
        exclude_sheets=tuple(str(item) for item in data.get("exclude_sheets", [])),
        sheet_map={
            str(expected): str(candidate)
            for expected, candidate in data.get("sheet_map", {}).items()
        },
        ranges=_parse_range_map(data.get("ranges"), "ranges"),
        ignore_ranges=_parse_range_map(
            data.get("ignore_ranges"), "ignore_ranges"
        ),
        sheet_weights=sheet_weights,
        range_weights=tuple(weighted_ranges),
        absolute_tolerance=absolute_tolerance,
        relative_tolerance=relative_tolerance,
        text_case_sensitive=bool(data.get("text_case_sensitive", False)),
        collapse_whitespace=bool(data.get("collapse_whitespace", True)),
        sheet_name_case_sensitive=bool(
            data.get("sheet_name_case_sensitive", False)
        ),
        formula_mode=formula_mode,  # type: ignore[arg-type]
        score_blank_expected_cells=bool(
            data.get("score_blank_expected_cells", False)
        ),
        penalize_extra_cells=bool(data.get("penalize_extra_cells", False)),
        penalize_extra_sheets=bool(data.get("penalize_extra_sheets", False)),
        pass_threshold=pass_threshold,
        max_mismatches=max_mismatches,
    )


def _normalized_sheet_name(name: str, config: WorkbookScoreConfig) -> str:
    return name if config.sheet_name_case_sensitive else name.casefold()


def _selected_expected_sheets(
    expected: WorkbookCells, config: WorkbookScoreConfig
) -> list[str]:
    include = {
        _normalized_sheet_name(name, config) for name in config.include_sheets
    }
    exclude = {
        _normalized_sheet_name(name, config) for name in config.exclude_sheets
    }
    result = []
    for sheet in expected:
        normalized = _normalized_sheet_name(sheet, config)
        if include and normalized not in include:
            continue
        if normalized in exclude:
            continue
        result.append(sheet)
    missing_includes = include - {
        _normalized_sheet_name(sheet, config) for sheet in result
    }
    if missing_includes:
        raise ValueError(
            f"include_sheets not found in expected workbook: {sorted(missing_includes)}"
        )
    return result


def _resolve_candidate_sheet(
    expected_sheet: str,
    candidate: WorkbookCells,
    config: WorkbookScoreConfig,
) -> str | None:
    requested = config.sheet_map.get(expected_sheet, expected_sheet)
    if requested in candidate:
        return requested
    if config.sheet_name_case_sensitive:
        return None
    matches = [
        name for name in candidate if name.casefold() == requested.casefold()
    ]
    if len(matches) > 1:
        raise ValueError(
            f"Candidate has ambiguous case-insensitive sheet names for {requested!r}"
        )
    return matches[0] if matches else None


def _inside_ranges(reference: str, ranges: tuple[CellRange, ...]) -> bool:
    if not ranges:
        return True
    row, col = _cell_coordinates(reference)
    return any(cell_range.contains(row, col) for cell_range in ranges)


def _inside_ignored(
    reference: str, sheet: str, config: WorkbookScoreConfig
) -> bool:
    ignored = config.ignore_ranges.get(sheet, ())
    if not ignored:
        return False
    row, col = _cell_coordinates(reference)
    return any(cell_range.contains(row, col) for cell_range in ignored)


def _is_blank(value: object) -> bool:
    return value is None or (isinstance(value, str) and value.strip() == "")


def _cell_has_content(cell: WorkbookCell | None) -> bool:
    return cell is not None and (cell.formula is not None or not _is_blank(cell.value))


def _expand_ranges(ranges: tuple[CellRange, ...], maximum: int = 1_000_000) -> set[str]:
    references: set[str] = set()
    for cell_range in ranges:
        size = (cell_range.end_row - cell_range.start_row + 1) * (
            cell_range.end_col - cell_range.start_col + 1
        )
        if len(references) + size > maximum:
            raise ValueError("Configured score ranges exceed one million cells")
        for row in range(cell_range.start_row, cell_range.end_row + 1):
            for col in range(cell_range.start_col, cell_range.end_col + 1):
                references.add(f"{_column_letters(col)}{row}")
    return references


def _expected_references(
    sheet: str,
    cells: dict[str, WorkbookCell],
    config: WorkbookScoreConfig,
) -> set[str]:
    ranges = config.ranges.get(sheet, ())
    if config.score_blank_expected_cells and ranges:
        candidates = _expand_ranges(ranges)
    else:
        candidates = set(cells)
    result: set[str] = set()
    for reference in candidates:
        if not _inside_ranges(reference, ranges):
            continue
        if _inside_ignored(reference, sheet, config):
            continue
        cell = cells.get(reference)
        if config.score_blank_expected_cells or _cell_has_content(cell):
            result.add(reference)
    return result


def _normalize_text(value: str, config: WorkbookScoreConfig) -> str:
    result = re.sub(r"\s+", " ", value).strip() if config.collapse_whitespace else value
    return result if config.text_case_sensitive else result.casefold()


def _value_matches(
    expected: object, candidate: object, config: WorkbookScoreConfig
) -> bool:
    if _is_blank(expected) or _is_blank(candidate):
        return _is_blank(expected) and _is_blank(candidate)
    if isinstance(expected, bool) or isinstance(candidate, bool):
        return type(expected) is type(candidate) and expected == candidate
    if isinstance(expected, (int, float)) and isinstance(candidate, (int, float)):
        return math.isfinite(float(candidate)) and math.isclose(
            float(expected),
            float(candidate),
            abs_tol=config.absolute_tolerance,
            rel_tol=config.relative_tolerance,
        )
    if isinstance(expected, (date, datetime)) and isinstance(
        candidate, (date, datetime)
    ):
        expected_date = expected.date() if isinstance(expected, datetime) else expected
        candidate_date = candidate.date() if isinstance(candidate, datetime) else candidate
        return expected_date == candidate_date
    return _normalize_text(str(expected), config) == _normalize_text(
        str(candidate), config
    )


def _normalize_formula(formula: str | None) -> str | None:
    if formula is None:
        return None
    return re.sub(r"\s+", "", formula).upper()


def _cell_matches(
    expected: WorkbookCell,
    candidate: WorkbookCell | None,
    config: WorkbookScoreConfig,
) -> bool:
    candidate = candidate or WorkbookCell(value="")
    value_match = _value_matches(expected.value, candidate.value, config)
    if config.formula_mode == "values":
        return value_match
    if expected.formula is None:
        return value_match
    formula_match = _normalize_formula(expected.formula) == _normalize_formula(
        candidate.formula
    )
    if config.formula_mode == "exact":
        return formula_match
    return value_match and formula_match


def _cell_weight(
    sheet: str, reference: str, config: WorkbookScoreConfig
) -> float:
    row, col = _cell_coordinates(reference)
    weight = config.sheet_weights.get(sheet, 1.0)
    for weighted_range in config.range_weights:
        if (
            weighted_range.sheet == sheet
            and weighted_range.cell_range.contains(row, col)
        ):
            weight *= weighted_range.weight
    return weight


def _format_cell(cell: WorkbookCell | None) -> str:
    if cell is None:
        return "<missing>"
    if cell.formula:
        return f"value={cell.value!r}, formula={cell.formula!r}"
    return repr(cell.value)


def score_workbook_cells(
    expected: WorkbookCells,
    candidate: WorkbookCells,
    config: WorkbookScoreConfig,
) -> WorkbookScoreReport:
    """Score two already-parsed workbooks."""
    expected_sheets = _selected_expected_sheets(expected, config)
    if not expected_sheets:
        raise ValueError("No expected worksheets are selected for scoring")

    resolved = {
        sheet: _resolve_candidate_sheet(sheet, candidate, config)
        for sheet in expected_sheets
    }
    matched_sheet_count = sum(name is not None for name in resolved.values())
    sheet_recall = matched_sheet_count / len(expected_sheets)
    mapped_candidate_sheets = {name for name in resolved.values() if name is not None}
    extra_sheets = tuple(
        name for name in candidate if name not in mapped_candidate_sheets
    )
    if config.penalize_extra_sheets:
        sheet_precision = matched_sheet_count / max(
            1, matched_sheet_count + len(extra_sheets)
        )
        structure_ratio = (
            2 * sheet_recall * sheet_precision / (sheet_recall + sheet_precision)
            if sheet_recall + sheet_precision
            else 0.0
        )
    else:
        structure_ratio = sheet_recall

    total_weight = 0.0
    covered_weight = 0.0
    matched_weight = 0.0
    expected_cell_count = 0
    covered_cell_count = 0
    matched_cell_count = 0
    extra_cell_count = 0
    extra_weight = 0.0
    sheet_scores: list[SheetScore] = []
    mismatches: list[str] = []

    for expected_sheet in expected_sheets:
        candidate_sheet = resolved[expected_sheet]
        expected_cells = expected[expected_sheet]
        actual_cells = candidate.get(candidate_sheet, {}) if candidate_sheet else {}
        references = _expected_references(expected_sheet, expected_cells, config)
        ranges = config.ranges.get(expected_sheet, ())
        sheet_total_weight = 0.0
        sheet_matched_weight = 0.0
        sheet_covered = 0
        sheet_matched = 0

        for reference in sorted(references, key=_cell_coordinates):
            expected_cell = expected_cells.get(reference, WorkbookCell(value=""))
            actual_cell = actual_cells.get(reference)
            weight = _cell_weight(expected_sheet, reference, config)
            total_weight += weight
            sheet_total_weight += weight
            expected_cell_count += 1
            if _cell_has_content(actual_cell) or _is_blank(expected_cell.value):
                covered_weight += weight
                covered_cell_count += 1
                sheet_covered += 1
            if _cell_matches(expected_cell, actual_cell, config):
                matched_weight += weight
                sheet_matched_weight += weight
                matched_cell_count += 1
                sheet_matched += 1
            elif len(mismatches) < config.max_mismatches:
                mismatch_type = "missing" if actual_cell is None else "different"
                mismatches.append(
                    f"{expected_sheet}!{reference} [{mismatch_type}]: "
                    f"expected {_format_cell(expected_cell)}, "
                    f"candidate {_format_cell(actual_cell)}"
                )

        sheet_extra = 0
        if candidate_sheet:
            for reference, actual_cell in actual_cells.items():
                if not _cell_has_content(actual_cell):
                    continue
                if not _inside_ranges(reference, ranges):
                    continue
                if _inside_ignored(reference, expected_sheet, config):
                    continue
                if reference not in references:
                    sheet_extra += 1
                    extra_weight += _cell_weight(
                        expected_sheet, reference, config
                    )
        extra_cell_count += sheet_extra
        sheet_scores.append(
            SheetScore(
                expected_sheet=expected_sheet,
                candidate_sheet=candidate_sheet,
                expected_cells=len(references),
                covered_cells=sheet_covered,
                matched_cells=sheet_matched,
                extra_cells=sheet_extra,
                weighted_accuracy=(
                    sheet_matched_weight / sheet_total_weight
                    if sheet_total_weight
                    else 1.0
                ),
            )
        )

    coverage_recall = covered_weight / total_weight if total_weight else 1.0
    if config.penalize_extra_cells:
        coverage_precision = covered_weight / max(
            1.0, covered_weight + extra_weight
        )
        coverage_ratio = (
            2
            * coverage_recall
            * coverage_precision
            / (coverage_recall + coverage_precision)
            if coverage_recall + coverage_precision
            else 0.0
        )
    else:
        coverage_ratio = coverage_recall
    value_ratio = matched_weight / total_weight if total_weight else 1.0
    structure_score = 10.0 * structure_ratio
    coverage_score = 20.0 * coverage_ratio
    value_score = 70.0 * value_ratio
    score = round(structure_score + coverage_score + value_score, 2)
    return WorkbookScoreReport(
        score=score,
        passed=score >= config.pass_threshold,
        structure_score=round(structure_score, 2),
        coverage_score=round(coverage_score, 2),
        value_score=round(value_score, 2),
        expected_sheets=len(expected_sheets),
        matched_sheets=matched_sheet_count,
        expected_cells=expected_cell_count,
        covered_cells=covered_cell_count,
        matched_cells=matched_cell_count,
        extra_cells=extra_cell_count,
        extra_sheets=extra_sheets,
        sheet_scores=tuple(sheet_scores),
        mismatches=tuple(mismatches),
    )


def score_workbooks(
    expected_path: Path,
    candidate_path: Path,
    config: WorkbookScoreConfig,
) -> WorkbookScoreReport:
    return score_workbook_cells(
        read_workbook_cells(expected_path),
        read_workbook_cells(candidate_path),
        config,
    )


def _print_report(report: WorkbookScoreReport) -> None:
    status = "PASS" if report.passed else "FAIL"
    print(f"[{status}] Generic workbook score: {report.score:.2f}/100")
    print(
        "  structure={:.2f}/10 coverage={:.2f}/20 values={:.2f}/70".format(
            report.structure_score,
            report.coverage_score,
            report.value_score,
        )
    )
    print(
        f"  sheets: {report.matched_sheets}/{report.expected_sheets} matched | "
        f"cells: {report.matched_cells}/{report.expected_cells} correct | "
        f"extras: {report.extra_cells}"
    )
    for sheet in report.sheet_scores:
        candidate_name = sheet.candidate_sheet or "<missing>"
        print(
            f"  {sheet.expected_sheet} -> {candidate_name}: "
            f"{100 * sheet.weighted_accuracy:.1f}% | "
            f"{sheet.matched_cells}/{sheet.expected_cells} cells"
        )
    if report.extra_sheets:
        print(f"  extra sheets: {', '.join(report.extra_sheets)}")
    if report.mismatches:
        print(f"  mismatches (first {len(report.mismatches)}):")
        for mismatch in report.mismatches:
            print(f"    {mismatch}")


def _build_parser() -> argparse.ArgumentParser:
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument("expected", type=Path, help="Gold/answer-key .xlsx workbook")
    parser.add_argument("candidate", type=Path, help="Candidate output .xlsx workbook")
    parser.add_argument("--config", type=Path, default=DEFAULT_CONFIG)
    parser.add_argument(
        "--json-report",
        type=Path,
        help="Optionally save the full score report as JSON",
    )
    return parser


def main() -> int:
    args = _build_parser().parse_args()
    report = score_workbooks(
        args.expected,
        args.candidate,
        load_workbook_score_config(args.config),
    )
    _print_report(report)
    if args.json_report:
        args.json_report.parent.mkdir(parents=True, exist_ok=True)
        args.json_report.write_text(
            json.dumps(report.to_dict(), indent=2) + "\n", encoding="utf-8"
        )
    return 0 if report.passed else 1


if __name__ == "__main__":
    raise SystemExit(main())

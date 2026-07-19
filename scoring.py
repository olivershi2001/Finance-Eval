"""Weighted, diagnostic scorer for customer-cube candidate rows."""

from __future__ import annotations

import math
from collections import Counter, defaultdict
from collections.abc import Iterable

from models import CubeConfig, CubeRow, EvalReport, SliceScore


MOVEMENT_FIELDS = (
    "opening_revenue",
    "new",
    "upsell",
    "downsell",
    "cross_sell",
    "churn",
    "closing_revenue",
)
RETENTION_FIELDS = (
    "gross_retention",
    "net_retention",
    "ttm_gross_retention",
)


def _matches(expected: float | None, actual: float | None, config: CubeConfig) -> bool:
    if expected is None or actual is None:
        return expected is None and actual is None
    if not math.isfinite(actual):
        return False
    return math.isclose(
        expected,
        actual,
        abs_tol=config.absolute_tolerance,
        rel_tol=config.relative_tolerance,
    )


def _row_reconciles(row: CubeRow, config: CubeConfig) -> bool:
    implied_closing = (
        row.opening_revenue
        + row.new
        + row.upsell
        + row.cross_sell
        - row.downsell
        - row.churn
    )
    if not _matches(row.closing_revenue, implied_closing, config):
        return False
    if row.opening_revenue <= 0:
        return row.gross_retention is None and row.net_retention is None
    implied_gross = max(
        0.0,
        min(
            1.0,
            (row.opening_revenue - row.downsell - row.churn)
            / row.opening_revenue,
        ),
    )
    implied_net = max(
        0.0,
        (
            row.opening_revenue
            + row.upsell
            + row.cross_sell
            - row.downsell
            - row.churn
        )
        / row.opening_revenue,
    )
    return _matches(row.gross_retention, implied_gross, config) and _matches(
        row.net_retention, implied_net, config
    )


def score_cube(
    expected_rows: Iterable[CubeRow],
    actual_rows: Iterable[CubeRow],
    config: CubeConfig,
) -> EvalReport:
    """Score schema/uniqueness, row coverage, values, and internal tie-outs."""
    expected = tuple(expected_rows)
    actual = tuple(actual_rows)
    counts = Counter(row.key for row in actual)
    duplicate_keys = tuple(sorted(key for key, count in counts.items() if count > 1))
    actual_map = {row.key: row for row in actual}
    expected_map = {row.key: row for row in expected}
    missing = tuple(sorted(set(expected_map) - set(actual_map)))
    extra = tuple(sorted(set(actual_map) - set(expected_map)))
    common = sorted(set(expected_map) & set(actual_map))

    valid_rows = sum(
        row.grain in ("quarterly", "annual")
        and row.slice_type in config.slices
        and row.slice_value != ""
        and all(
            math.isfinite(value)
            for value in (
                row.opening_revenue,
                row.new,
                row.upsell,
                row.downsell,
                row.cross_sell,
                row.churn,
                row.closing_revenue,
            )
        )
        for row in actual
    )
    row_validity = valid_rows / len(actual) if actual else 0.0
    uniqueness = 1.0 - len(duplicate_keys) / max(1, len(expected))
    structure_score = 10.0 * max(0.0, row_validity * uniqueness)

    recall = len(common) / len(expected) if expected else 1.0
    precision = len(common) / len(actual_map) if actual_map else 0.0
    coverage_score = 20.0 * (2 * recall * precision / (recall + precision) if recall + precision else 0.0)

    movement_matches = 0
    movement_total = len(expected) * len(MOVEMENT_FIELDS)
    retention_matches = 0
    retention_total = len(expected) * len(RETENTION_FIELDS)
    mismatches: list[str] = []
    grouped_matches: dict[tuple[str, str], list[int]] = defaultdict(lambda: [0, 0])

    for key in common:
        expected_row = expected_map[key]
        actual_row = actual_map[key]
        group = (expected_row.grain, expected_row.slice_type)
        for field in MOVEMENT_FIELDS:
            matched = _matches(
                getattr(expected_row, field), getattr(actual_row, field), config
            )
            movement_matches += int(matched)
            grouped_matches[group][0] += int(matched)
            grouped_matches[group][1] += 1
            if not matched and len(mismatches) < 20:
                mismatches.append(
                    f"{key} {field}: expected={getattr(expected_row, field)!r}, "
                    f"actual={getattr(actual_row, field)!r}"
                )
        for field in RETENTION_FIELDS:
            matched = _matches(
                getattr(expected_row, field), getattr(actual_row, field), config
            )
            retention_matches += int(matched)
            grouped_matches[group][0] += int(matched)
            grouped_matches[group][1] += 1
            if not matched and len(mismatches) < 20:
                mismatches.append(
                    f"{key} {field}: expected={getattr(expected_row, field)!r}, "
                    f"actual={getattr(actual_row, field)!r}"
                )

    movement_score = 40.0 * movement_matches / movement_total if movement_total else 40.0
    retention_score = 20.0 * retention_matches / retention_total if retention_total else 20.0
    reconciliation_score = 10.0 * (
        sum(_row_reconciles(row, config) for row in actual_map.values())
        / len(actual_map)
        if actual_map
        else 0.0
    )
    score = round(
        structure_score
        + coverage_score
        + movement_score
        + retention_score
        + reconciliation_score,
        2,
    )
    slice_scores = tuple(
        SliceScore(
            grain=group[0],
            slice_type=group[1],
            matched_cells=values[0],
            total_cells=values[1],
        )
        for group, values in sorted(grouped_matches.items())
    )
    return EvalReport(
        score=score,
        passed=score >= config.pass_threshold,
        structure_score=round(structure_score, 2),
        coverage_score=round(coverage_score, 2),
        movement_score=round(movement_score, 2),
        retention_score=round(retention_score, 2),
        reconciliation_score=round(reconciliation_score, 2),
        expected_rows=len(expected),
        actual_rows=len(actual),
        missing_keys=missing[:20],
        extra_keys=extra[:20],
        duplicate_keys=duplicate_keys[:20],
        mismatches=tuple(mismatches),
        slice_scores=slice_scores,
    )

"""Typed contracts for the customer-cube evaluation harness."""

from __future__ import annotations

from dataclasses import dataclass, field
from datetime import date
from typing import Literal


Grain = Literal["quarterly", "annual"]
SliceType = Literal["overall", "product", "customer"]
AmountMode = Literal["period_sum", "period_end"]


@dataclass(frozen=True, slots=True)
class RevenueRecord:
    """One normalized row from the four-column raw workbook."""

    customer_name: str
    month: date
    amount: float
    product: str


@dataclass(frozen=True, slots=True)
class CubeConfig:
    """Business conventions that affect the expected cube."""

    amount_mode: AmountMode = "period_sum"
    fiscal_year_start_month: int = 1
    slices: tuple[SliceType, ...] = ("overall", "product", "customer")
    absolute_tolerance: float = 0.01
    relative_tolerance: float = 1e-6
    pass_threshold: float = 95.0

    def __post_init__(self) -> None:
        if not 1 <= self.fiscal_year_start_month <= 12:
            raise ValueError("fiscal_year_start_month must be between 1 and 12")
        if self.amount_mode not in ("period_sum", "period_end"):
            raise ValueError(f"Unsupported amount_mode: {self.amount_mode!r}")
        invalid_slices = set(self.slices) - {"overall", "product", "customer"}
        if invalid_slices:
            raise ValueError(f"Unsupported slice(s): {sorted(invalid_slices)!r}")


@dataclass(frozen=True, slots=True)
class CubeRow:
    """One period/slice result returned by a candidate."""

    grain: Grain
    period: str
    slice_type: SliceType
    slice_value: str
    opening_revenue: float
    new: float
    upsell: float
    downsell: float
    cross_sell: float
    churn: float
    closing_revenue: float
    gross_retention: float | None
    net_retention: float | None
    ttm_gross_retention: float | None

    @property
    def key(self) -> tuple[str, str, str, str]:
        return (self.grain, self.period, self.slice_type, self.slice_value)

    def to_dict(self) -> dict[str, str | float | None]:
        return {
            "grain": self.grain,
            "period": self.period,
            "slice_type": self.slice_type,
            "slice_value": self.slice_value,
            "opening_revenue": self.opening_revenue,
            "new": self.new,
            "upsell": self.upsell,
            "downsell": self.downsell,
            "cross_sell": self.cross_sell,
            "churn": self.churn,
            "closing_revenue": self.closing_revenue,
            "gross_retention": self.gross_retention,
            "net_retention": self.net_retention,
            "ttm_gross_retention": self.ttm_gross_retention,
        }

    @classmethod
    def from_dict(cls, data: dict[str, object]) -> "CubeRow":
        def number(name: str, *, nullable: bool = False) -> float | None:
            value = data.get(name)
            if value is None and nullable:
                return None
            if isinstance(value, bool) or not isinstance(value, (int, float)):
                raise ValueError(f"{name} must be numeric")
            return float(value)

        grain = str(data["grain"])
        slice_type = str(data["slice_type"])
        if grain not in ("quarterly", "annual"):
            raise ValueError(f"Invalid grain: {grain!r}")
        if slice_type not in ("overall", "product", "customer"):
            raise ValueError(f"Invalid slice_type: {slice_type!r}")
        return cls(
            grain=grain,  # type: ignore[arg-type]
            period=str(data["period"]),
            slice_type=slice_type,  # type: ignore[arg-type]
            slice_value=str(data["slice_value"]),
            opening_revenue=float(number("opening_revenue")),
            new=float(number("new")),
            upsell=float(number("upsell")),
            downsell=float(number("downsell")),
            cross_sell=float(number("cross_sell")),
            churn=float(number("churn")),
            closing_revenue=float(number("closing_revenue")),
            gross_retention=number("gross_retention", nullable=True),
            net_retention=number("net_retention", nullable=True),
            ttm_gross_retention=number(
                "ttm_gross_retention", nullable=True
            ),
        )


@dataclass(frozen=True, slots=True)
class CandidateResult:
    """Swappable candidate contract."""

    rows: tuple[CubeRow, ...]


@dataclass(frozen=True, slots=True)
class SliceScore:
    grain: str
    slice_type: str
    matched_cells: int
    total_cells: int

    @property
    def accuracy(self) -> float:
        return self.matched_cells / self.total_cells if self.total_cells else 1.0


@dataclass(frozen=True, slots=True)
class EvalReport:
    score: float
    passed: bool
    structure_score: float
    coverage_score: float
    movement_score: float
    retention_score: float
    reconciliation_score: float
    expected_rows: int
    actual_rows: int
    missing_keys: tuple[tuple[str, str, str, str], ...] = ()
    extra_keys: tuple[tuple[str, str, str, str], ...] = ()
    duplicate_keys: tuple[tuple[str, str, str, str], ...] = ()
    mismatches: tuple[str, ...] = ()
    slice_scores: tuple[SliceScore, ...] = ()
    errors: tuple[str, ...] = field(default_factory=tuple)

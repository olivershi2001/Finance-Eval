"""Deterministic reference implementation of customer-cube SaaS metrics."""

from __future__ import annotations

from collections import defaultdict
from collections.abc import Iterable, Mapping
from datetime import date

from models import CubeConfig, CubeRow, Grain, RevenueRecord, SliceType


RevenueMap = dict[tuple[str, str], float]


def _month_index(value: date) -> int:
    return value.year * 12 + value.month - 1


def _fiscal_year(value: date, start_month: int) -> int:
    return value.year + (1 if value.month >= start_month and start_month != 1 else 0)


def _quarter_key(value: date, start_month: int) -> tuple[int, int]:
    fiscal_year = _fiscal_year(value, start_month)
    fiscal_month = (value.month - start_month) % 12
    return fiscal_year, fiscal_month // 3 + 1


def _period_label(grain: Grain, key: tuple[int, int] | int) -> str:
    if grain == "quarterly":
        year, quarter = key  # type: ignore[misc]
        return f"FY{year}-Q{quarter}"
    return f"FY{key}"


def _aggregate_periods(
    records: Iterable[RevenueRecord],
    grain: Grain,
    config: CubeConfig,
) -> tuple[list[tuple[int, int] | int], dict[tuple[int, int] | int, RevenueMap]]:
    grouped: dict[tuple[int, int] | int, RevenueMap] = defaultdict(
        lambda: defaultdict(float)
    )
    period_months: dict[tuple[int, int] | int, dict[tuple[str, str], dict[int, float]]] = (
        defaultdict(lambda: defaultdict(lambda: defaultdict(float)))
    )

    materialized = tuple(records)
    for record in materialized:
        if record.amount == 0:
            continue
        key: tuple[int, int] | int
        if grain == "quarterly":
            key = _quarter_key(record.month, config.fiscal_year_start_month)
        else:
            key = _fiscal_year(record.month, config.fiscal_year_start_month)
        customer_product = (record.customer_name, record.product)
        grouped[key][customer_product] += record.amount
        period_months[key][customer_product][_month_index(record.month)] += record.amount

    if materialized:
        first_month = min(_month_index(record.month) for record in materialized)
        last_month = max(_month_index(record.month) for record in materialized)
        for month_index in range(first_month, last_month + 1):
            month_date = date(month_index // 12, month_index % 12 + 1, 1)
            period_key: tuple[int, int] | int
            if grain == "quarterly":
                period_key = _quarter_key(month_date, config.fiscal_year_start_month)
            else:
                period_key = _fiscal_year(month_date, config.fiscal_year_start_month)
            grouped.setdefault(period_key, defaultdict(float))
            period_months.setdefault(
                period_key, defaultdict(lambda: defaultdict(float))
            )

    if config.amount_mode == "period_end":
        end_balances: dict[tuple[int, int] | int, RevenueMap] = defaultdict(
            lambda: defaultdict(float)
        )
        for period, customer_products in period_months.items():
            latest_month = _period_end_month_index(
                grain, period, config.fiscal_year_start_month
            )
            for customer_product, monthly_values in customer_products.items():
                value = monthly_values.get(latest_month, 0.0)
                if value > 0:
                    end_balances[period][customer_product] = value
            end_balances.setdefault(period, defaultdict(float))
        grouped = end_balances

    return sorted(grouped), grouped


def _customer_totals(revenue: Mapping[tuple[str, str], float]) -> dict[str, float]:
    totals: dict[str, float] = defaultdict(float)
    for (customer, _product), amount in revenue.items():
        totals[customer] += amount
    return totals


def _movement_components(
    opening: Mapping[tuple[str, str], float],
    closing: Mapping[tuple[str, str], float],
) -> dict[tuple[str, str], dict[str, float]]:
    """Classify product-level movement using customer-level new/churn status."""
    opening_customer = _customer_totals(opening)
    closing_customer = _customer_totals(closing)
    result: dict[tuple[str, str], dict[str, float]] = {}

    for customer_product in sorted(set(opening) | set(closing)):
        customer, _product = customer_product
        prior = max(0.0, float(opening.get(customer_product, 0.0)))
        current = max(0.0, float(closing.get(customer_product, 0.0)))
        metrics = {
            "opening_revenue": prior,
            "new": 0.0,
            "upsell": 0.0,
            "downsell": 0.0,
            "cross_sell": 0.0,
            "churn": 0.0,
            "closing_revenue": current,
        }

        if opening_customer.get(customer, 0.0) <= 0 and current > 0:
            metrics["new"] = current
        elif closing_customer.get(customer, 0.0) <= 0 and prior > 0:
            metrics["churn"] = prior
        elif prior <= 0 < current:
            metrics["cross_sell"] = current
        elif current > prior:
            metrics["upsell"] = current - prior
        elif prior > current:
            metrics["downsell"] = prior - current

        result[customer_product] = metrics

    return result


def _slice_values(
    movements: Mapping[tuple[str, str], Mapping[str, float]],
    slice_type: SliceType,
) -> list[str]:
    if slice_type == "overall":
        return ["All"]
    index = 1 if slice_type == "product" else 0
    return sorted({key[index] for key in movements})


def _included(
    customer_product: tuple[str, str],
    slice_type: SliceType,
    slice_value: str,
) -> bool:
    if slice_type == "overall":
        return True
    if slice_type == "product":
        return customer_product[1] == slice_value
    return customer_product[0] == slice_value


def _sum_metrics(
    movements: Mapping[tuple[str, str], Mapping[str, float]],
    slice_type: SliceType,
    slice_value: str,
) -> dict[str, float]:
    totals = {
        "opening_revenue": 0.0,
        "new": 0.0,
        "upsell": 0.0,
        "downsell": 0.0,
        "cross_sell": 0.0,
        "churn": 0.0,
        "closing_revenue": 0.0,
    }
    for customer_product, metrics in movements.items():
        if _included(customer_product, slice_type, slice_value):
            for metric in totals:
                totals[metric] += metrics[metric]
    return totals


def _retention(totals: Mapping[str, float]) -> tuple[float | None, float | None]:
    opening = totals["opening_revenue"]
    if opening <= 0:
        return None, None
    gross = (opening - totals["downsell"] - totals["churn"]) / opening
    net = (
        opening
        + totals["upsell"]
        + totals["cross_sell"]
        - totals["downsell"]
        - totals["churn"]
    ) / opening
    return max(0.0, min(1.0, gross)), max(0.0, net)


def _aggregate_month_window(
    records: Iterable[RevenueRecord],
    start_index: int,
    end_index: int,
) -> RevenueMap:
    result: RevenueMap = defaultdict(float)
    for record in records:
        month = _month_index(record.month)
        if start_index <= month <= end_index:
            result[(record.customer_name, record.product)] += record.amount
    return result


def _ttm_gross_by_slice(
    records: tuple[RevenueRecord, ...],
    period_end_index: int,
    slice_type: SliceType,
    slice_value: str,
    config: CubeConfig,
) -> float | None:
    earliest = min((_month_index(record.month) for record in records), default=0)
    if config.amount_mode == "period_end":
        prior_month = period_end_index - 12
        if prior_month < earliest:
            return None
        prior = _aggregate_month_window(records, prior_month, prior_month)
        current = _aggregate_month_window(records, period_end_index, period_end_index)
    else:
        current_start = period_end_index - 11
        prior_start = period_end_index - 23
        if prior_start < earliest:
            return None
        prior = _aggregate_month_window(records, prior_start, current_start - 1)
        current = _aggregate_month_window(records, current_start, period_end_index)
    totals = _sum_metrics(
        _movement_components(prior, current), slice_type, slice_value
    )
    gross, _net = _retention(totals)
    return gross


def _period_end_month_index(
    grain: Grain,
    key: tuple[int, int] | int,
    fiscal_year_start_month: int,
) -> int:
    if grain == "annual":
        fiscal_year = int(key)  # type: ignore[arg-type]
        end_month = 12 if fiscal_year_start_month == 1 else fiscal_year_start_month - 1
        end_year = fiscal_year if fiscal_year_start_month == 1 else fiscal_year
        return end_year * 12 + end_month - 1
    fiscal_year, quarter = key  # type: ignore[misc]
    start_year = fiscal_year if fiscal_year_start_month == 1 else fiscal_year - 1
    start_index = start_year * 12 + fiscal_year_start_month - 1
    return start_index + quarter * 3 - 1


def _calculate_grain(
    records: tuple[RevenueRecord, ...],
    grain: Grain,
    config: CubeConfig,
) -> list[CubeRow]:
    periods, revenue_by_period = _aggregate_periods(records, grain, config)
    rows: list[CubeRow] = []

    for index, period in enumerate(periods):
        opening = revenue_by_period[periods[index - 1]] if index else {}
        closing = revenue_by_period[period]
        movements = _movement_components(opening, closing)
        period_end = _period_end_month_index(
            grain, period, config.fiscal_year_start_month
        )

        for slice_type in config.slices:
            for slice_value in _slice_values(movements, slice_type):
                totals = _sum_metrics(movements, slice_type, slice_value)
                gross, net = _retention(totals)
                ttm_gross = _ttm_gross_by_slice(
                    records, period_end, slice_type, slice_value, config
                )
                rows.append(
                    CubeRow(
                        grain=grain,
                        period=_period_label(grain, period),
                        slice_type=slice_type,
                        slice_value=slice_value,
                        opening_revenue=round(totals["opening_revenue"], 10),
                        new=round(totals["new"], 10),
                        upsell=round(totals["upsell"], 10),
                        downsell=round(totals["downsell"], 10),
                        cross_sell=round(totals["cross_sell"], 10),
                        churn=round(totals["churn"], 10),
                        closing_revenue=round(totals["closing_revenue"], 10),
                        gross_retention=None if gross is None else round(gross, 10),
                        net_retention=None if net is None else round(net, 10),
                        ttm_gross_retention=(
                            None if ttm_gross is None else round(ttm_gross, 10)
                        ),
                    )
                )
    return rows


def calculate_cube(
    records: tuple[RevenueRecord, ...],
    config: CubeConfig,
) -> tuple[CubeRow, ...]:
    """Calculate quarterly and annual customer-cube outputs."""
    if any(record.amount < 0 for record in records):
        raise ValueError("Amount must be non-negative for movement classification")
    if any(not record.customer_name.strip() or not record.product.strip() for record in records):
        raise ValueError("Customer name and product cannot be blank")
    normalized = tuple(
        RevenueRecord(
            customer_name=record.customer_name.strip(),
            month=date(record.month.year, record.month.month, 1),
            amount=float(record.amount),
            product=record.product.strip(),
        )
        for record in records
    )
    return tuple(
        _calculate_grain(normalized, "quarterly", config)
        + _calculate_grain(normalized, "annual", config)
    )

"""Reference candidate for the customer-cube eval.

Replace ``build_customer_cube`` with the model, agent, or implementation under
test. The eval runner and frozen expected results should not need to change.
"""

from __future__ import annotations

from customer_cube import calculate_cube
from models import CandidateResult, CubeConfig, RevenueRecord


def build_customer_cube(
    records: tuple[RevenueRecord, ...],
    config: CubeConfig,
) -> CandidateResult:
    """Return quarterly and annual rows for all configured cuts."""
    return CandidateResult(rows=calculate_cube(records, config))

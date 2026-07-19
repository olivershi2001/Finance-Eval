"""Regenerate the frozen gold JSON from the deterministic reference engine."""

from __future__ import annotations

import sys
from pathlib import Path


ROOT = Path(__file__).resolve().parents[1]
sys.path.insert(0, str(ROOT))

from candidate import build_customer_cube  # noqa: E402
from eval_runner import load_config, save_cube_rows  # noqa: E402
from xlsx_reader import load_revenue_records  # noqa: E402


def main() -> None:
    records = load_revenue_records(ROOT / "fixtures" / "customer_cube_raw.xlsx")
    rows = build_customer_cube(records, load_config()).rows
    save_cube_rows(ROOT / "scenarios" / "customer_cube_expected.json", rows)
    print(f"Wrote {len(rows)} frozen gold rows")


if __name__ == "__main__":
    main()

"""Clean CSV extracts for the Tableau dashboard."""

from __future__ import annotations

from pathlib import Path

import pandas as pd


def write_tableau_feed(out_dir: Path, sales: pd.DataFrame, quotas: pd.DataFrame) -> list[Path]:
    """Write fact_sales.csv (one row per order line) and region_month_summary.csv (with quota)."""
    out_dir.mkdir(parents=True, exist_ok=True)

    fact = sales.copy()
    fact["year"] = fact["order_date"].dt.year
    fact["quarter"] = fact["year"].astype(str) + " Q" + fact["order_date"].dt.quarter.astype(str)
    fact_path = out_dir / "fact_sales.csv"
    fact.to_csv(fact_path, index=False, date_format="%Y-%m-%d")

    monthly = (sales.groupby(["region", "month"], as_index=False)
               .agg(revenue=("sales", "sum"), orders=("order_count", "sum")))
    summary = monthly.merge(quotas[["region", "month", "py_revenue", "quota"]], on=["region", "month"], how="left")
    summary_path = out_dir / "region_month_summary.csv"
    summary.to_csv(summary_path, index=False, date_format="%Y-%m-%d")
    return [fact_path, summary_path]

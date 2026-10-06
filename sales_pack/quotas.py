"""Monthly quotas by region."""

from __future__ import annotations

import logging
from pathlib import Path

import pandas as pd

log = logging.getLogger(__name__)


def build_quotas(sales: pd.DataFrame, growth: float, override_path: str | Path | None = None) -> pd.DataFrame:
    """Return one row per region and month with columns region, month, py_revenue, quota, source.

    The default quota is the same month last year x (1 + growth), a common way to set targets
    when finance quotas aren't available. An override CSV (columns: region, month, quota)
    replaces the derived number wherever it has a value.
    """
    monthly = sales.groupby(["region", "month"], as_index=False)["sales"].sum()
    first, last = monthly["month"].min(), monthly["month"].max()
    months = pd.date_range(first + pd.DateOffset(years=1), pd.Timestamp(last.year, 12, 1), freq="MS")
    regions = sorted(sales["region"].unique())
    grid = pd.MultiIndex.from_product([regions, months], names=["region", "month"]).to_frame(index=False)

    prior = monthly.assign(month=monthly["month"] + pd.DateOffset(years=1)).rename(columns={"sales": "py_revenue"})
    grid = grid.merge(prior, on=["region", "month"], how="left")
    grid["py_revenue"] = grid["py_revenue"].fillna(0.0)
    grid["quota"] = grid["py_revenue"] * (1 + growth)
    grid["source"] = "derived"

    if override_path:
        override = _read_override(Path(override_path))
        grid = grid.merge(override, on=["region", "month"], how="left")
        has_value = grid["quota_override"].notna()
        grid.loc[has_value, "quota"] = grid.loc[has_value, "quota_override"]
        grid.loc[has_value, "source"] = "override"
        grid = grid.drop(columns="quota_override")
        log.info("Quota override applied to %d of %d region-months", int(has_value.sum()), len(grid))

    if grid.empty:
        log.warning("Less than 13 months of data, so quotas (and attainment) can't be derived.")
    return grid


def _read_override(path: Path) -> pd.DataFrame:
    if not path.exists():
        raise FileNotFoundError(f"Quota file not found: {path}")
    df = pd.read_csv(path)
    df.columns = [str(c).strip().lower() for c in df.columns]
    missing = {"region", "month", "quota"} - set(df.columns)
    if missing:
        raise ValueError(f"Quota file {path} is missing column(s): {', '.join(sorted(missing))}")
    df["region"] = df["region"].astype(str).str.strip()
    df["month"] = pd.to_datetime(df["month"], format="mixed").dt.to_period("M").dt.to_timestamp()
    df["quota_override"] = pd.to_numeric(df["quota"], errors="coerce")
    return df[["region", "month", "quota_override"]]

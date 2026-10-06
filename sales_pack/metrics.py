"""Pandas versions of the pack's metrics.

They drive the written callouts and the tests. The Excel workbook computes the same numbers
with formulas, so the two can be cross-checked against each other.
"""

from __future__ import annotations

import numpy as np
import pandas as pd

from .config import PY_WEEK_OFFSET_DAYS
from .periods import Period


def window(sales: pd.DataFrame, start, end) -> pd.DataFrame:
    """Rows with an order date between start and end, inclusive."""
    return sales[(sales["order_date"] >= start) & (sales["order_date"] <= end)]


def safe_div(numerator, denominator):
    """Division that returns NaN instead of inf when the denominator is zero."""
    if isinstance(denominator, pd.Series):
        return numerator / denominator.replace(0, np.nan)
    return numerator / denominator if denominator else np.nan


def quota_to_date(quotas: pd.DataFrame, start, end) -> pd.Series:
    """Quota for [start, end] by region, pro-rating each month by the days that fall in the window."""
    if quotas.empty:
        return pd.Series(dtype=float)
    month_start = quotas["month"]
    days_in_month = month_start.dt.days_in_month
    month_end = month_start + pd.to_timedelta(days_in_month - 1, unit="D")
    overlap_start = month_start.where(month_start > start, start)  # later of month start and window start
    overlap_end = month_end.where(month_end < end, end)            # earlier of month end and window end
    overlap_days = ((overlap_end - overlap_start).dt.days + 1).clip(lower=0)
    prorated = quotas["quota"] * overlap_days / days_in_month
    return prorated.groupby(quotas["region"]).sum()


def regional_table(sales: pd.DataFrame, quotas: pd.DataFrame, period: Period) -> pd.DataFrame:
    """Revenue, quota to date, attainment, orders, AOV and growth vs last year by region (+ Total)."""
    regions = sorted(sales["region"].unique())
    current = window(sales, period.start, period.end)
    prior = window(sales, period.py_start, period.py_end)

    table = pd.DataFrame(index=pd.Index(regions, name="region"))
    table["revenue"] = current.groupby("region")["sales"].sum()
    table["quota_to_date"] = quota_to_date(quotas, period.start, period.end)
    table["orders"] = current.groupby("region")["order_count"].sum()
    table["py_revenue"] = prior.groupby("region")["sales"].sum()
    table["py_orders"] = prior.groupby("region")["order_count"].sum()
    table = table.fillna(0.0)
    table.loc["Total"] = table.sum()

    table["attainment"] = safe_div(table["revenue"], table["quota_to_date"])
    table["gap_to_quota"] = table["revenue"] - table["quota_to_date"]
    table["aov"] = safe_div(table["revenue"], table["orders"])
    table["py_aov"] = safe_div(table["py_revenue"], table["py_orders"])
    table["growth_vs_py"] = safe_div(table["revenue"], table["py_revenue"]) - 1
    return table


def weekly_trend(sales: pd.DataFrame, as_of, weeks: int) -> pd.DataFrame:
    """Revenue by region for the last `weeks` weeks, plus the same week last year (52 weeks earlier)."""
    regions = sorted(sales["region"].unique())
    offset = pd.Timedelta(days=PY_WEEK_OFFSET_DAYS)
    rows = []
    for k in range(weeks - 1, -1, -1):
        end = pd.Timestamp(as_of) - pd.Timedelta(days=7 * k)
        start = end - pd.Timedelta(days=6)
        current = window(sales, start, end)
        prior = window(sales, start - offset, end - offset)
        by_region = current.groupby("region")["sales"].sum()
        row = {"week_ending": end}
        row.update({r: float(by_region.get(r, 0.0)) for r in regions})
        row["total"] = float(current["sales"].sum())
        row["py_total"] = float(prior["sales"].sum())
        rows.append(row)
    trend = pd.DataFrame(rows)
    trend["growth_vs_py"] = safe_div(trend["total"], trend["py_total"]) - 1
    return trend


def category_mix(sales: pd.DataFrame, period: Period) -> pd.DataFrame:
    """Share of revenue by category vs the same dates last year."""
    current = window(sales, period.start, period.end)
    prior = window(sales, period.py_start, period.py_end)
    categories = sorted(sales["category"].unique())
    mix = pd.DataFrame(index=pd.Index(categories, name="category"))
    mix["revenue"] = current.groupby("category")["sales"].sum()
    mix["py_revenue"] = prior.groupby("category")["sales"].sum()
    mix = mix.fillna(0.0)
    mix["share"] = safe_div(mix["revenue"], mix["revenue"].sum())
    mix["py_share"] = safe_div(mix["py_revenue"], mix["py_revenue"].sum())
    mix["mix_shift_pts"] = (mix["share"] - mix["py_share"]) * 100
    mix["growth_vs_py"] = safe_div(mix["revenue"], mix["py_revenue"]) - 1
    return mix.sort_values("revenue", ascending=False)


def subcategory_drivers(sales: pd.DataFrame, period: Period) -> pd.DataFrame:
    """Revenue by sub-category vs last year, sorted by current revenue."""
    keys = ["sub_category", "category"]
    current = window(sales, period.start, period.end).groupby(keys)["sales"].sum().rename("revenue")
    prior = window(sales, period.py_start, period.py_end).groupby(keys)["sales"].sum().rename("py_revenue")
    drivers = sales[keys].drop_duplicates().set_index(keys)
    drivers = drivers.join(current).join(prior).fillna(0.0).reset_index()
    drivers["change"] = drivers["revenue"] - drivers["py_revenue"]
    drivers["growth_vs_py"] = safe_div(drivers["revenue"], drivers["py_revenue"]) - 1
    return drivers.sort_values(["revenue", "sub_category"], ascending=[False, True]).reset_index(drop=True)

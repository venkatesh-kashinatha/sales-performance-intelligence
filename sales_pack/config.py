"""Project-wide settings and column aliases."""

from __future__ import annotations

# Raw exports name columns differently, so each standard field accepts a few aliases
# (matched case-insensitively, ignoring surrounding spaces).
REQUIRED_COLUMNS: dict[str, tuple[str, ...]] = {
    "order_id": ("order id", "order_id", "orderid", "order number"),
    "order_date": ("order date", "order_date", "orderdate", "date"),
    "region": ("region",),
    "category": ("category", "product category"),
    "sub_category": ("sub-category", "sub_category", "subcategory", "sub category"),
    "sales": ("sales", "revenue", "sales amount", "amount"),
}

OPTIONAL_COLUMNS: dict[str, tuple[str, ...]] = {
    "state": ("state", "state/province", "province"),
    "segment": ("segment", "customer segment"),
    "quantity": ("quantity", "qty", "units"),
    "profit": ("profit",),
}

# Column order of the cleaned table (and of the Data sheet / Tableau feed).
CLEAN_COLUMNS = [
    "order_id", "order_date", "week_ending", "month", "region", "state", "segment",
    "category", "sub_category", "sales", "quantity", "profit", "order_count",
]

DEFAULT_GROWTH = 0.10      # quota = same month last year x (1 + growth), unless a quota file is given
TREND_WEEKS = 13           # weeks shown on the Weekly Trend sheet
PY_WEEK_OFFSET_DAYS = 364  # 52 weeks back keeps weekdays aligned for week-vs-last-year comparisons

"""Reporting windows for the weekly review."""

from __future__ import annotations

from dataclasses import dataclass

import pandas as pd

from .config import PY_WEEK_OFFSET_DAYS

# Excel defined name that holds each window's start date (see the Summary sheet).
START_NAMES = {"week": "WeekStart", "mtd": "MonthStart", "qtd": "QtrStart", "ytd": "YearStart"}


@dataclass(frozen=True)
class Period:
    key: str    # week, mtd, qtd or ytd
    label: str
    start: pd.Timestamp
    end: pd.Timestamp

    # Last year's comparison window. The week goes back 52 weeks so weekdays line up (retail
    # convention); month, quarter and year to date use the same calendar dates last year.
    @property
    def py_start(self) -> pd.Timestamp:
        if self.key == "week":
            return self.start - pd.Timedelta(days=PY_WEEK_OFFSET_DAYS)
        return same_day_last_year(self.start)

    @property
    def py_end(self) -> pd.Timestamp:
        if self.key == "week":
            return self.end - pd.Timedelta(days=PY_WEEK_OFFSET_DAYS)
        return same_day_last_year(self.end)

    @property
    def excel_start_name(self) -> str:
        return START_NAMES[self.key]

    @property
    def excel_py_window(self) -> tuple[str, str]:
        """Excel expressions for the start and end of last year's comparison window."""
        start = self.excel_start_name
        if self.key == "week":
            return f"({start}-{PY_WEEK_OFFSET_DAYS})", f"(AsOf-{PY_WEEK_OFFSET_DAYS})"
        return f"EDATE({start},-12)", "EDATE(AsOf,-12)"


def same_day_last_year(ts) -> pd.Timestamp:
    """Same calendar day one year earlier. Matches Excel's EDATE(date, -12): Feb 29 maps to Feb 28."""
    return pd.Timestamp(ts) - pd.DateOffset(years=1)


def default_as_of(last_order_date) -> pd.Timestamp:
    """Most recent Sunday on or before the last order date, i.e. the last complete week."""
    d = pd.Timestamp(last_order_date).normalize()
    return d - pd.Timedelta(days=(d.weekday() + 1) % 7)


def build_periods(as_of) -> list[Period]:
    as_of = pd.Timestamp(as_of).normalize()
    quarter_start = pd.Timestamp(as_of.year, 3 * ((as_of.month - 1) // 3) + 1, 1)
    return [
        Period("week", "Week", as_of - pd.Timedelta(days=6), as_of),
        Period("mtd", "Month to date", as_of.replace(day=1), as_of),
        Period("qtd", "Quarter to date", quarter_start, as_of),
        Period("ytd", "Year to date", pd.Timestamp(as_of.year, 1, 1), as_of),
    ]

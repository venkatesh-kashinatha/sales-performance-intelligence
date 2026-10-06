"""Load a raw sales extract (CSV or Excel) and return a clean, standardised table."""

from __future__ import annotations

import logging
from dataclasses import dataclass
from pathlib import Path

import pandas as pd

from .config import CLEAN_COLUMNS, OPTIONAL_COLUMNS, REQUIRED_COLUMNS

log = logging.getLogger(__name__)


@dataclass
class LoadResult:
    data: pd.DataFrame
    source: Path
    rows_read: int
    rows_dropped: int


def load_sales(path: str | Path) -> LoadResult:
    """Read the extract, standardise column names and types, and add helper fields."""
    path = Path(path)
    if not path.exists():
        raise FileNotFoundError(f"Input file not found: {path}")

    raw = _read_raw(path)
    mapping = _match_columns(raw.columns)
    df = raw[list(mapping)].rename(columns=mapping)
    rows_read = len(df)

    df["order_date"] = pd.to_datetime(df["order_date"], format="mixed", errors="coerce").dt.normalize()
    df["sales"] = pd.to_numeric(
        df["sales"].astype("string").str.replace(r"[$,\s]", "", regex=True), errors="coerce"
    )
    for col in ("order_id", "region", "category", "sub_category"):
        df[col] = df[col].astype("string").str.strip().replace("", pd.NA)

    df = df.dropna(subset=["order_date", "sales", "order_id", "region", "category", "sub_category"])
    rows_dropped = rows_read - len(df)
    if df.empty:
        raise ValueError("No usable rows after cleaning. Check the date, sales and region columns.")

    for col in ("order_id", "region", "category", "sub_category"):
        df[col] = df[col].astype(str)
    for col in ("state", "segment"):
        df[col] = df[col].astype("string").str.strip().fillna("").astype(str) if col in df else ""
    for col in ("quantity", "profit"):
        df[col] = pd.to_numeric(df[col], errors="coerce") if col in df else float("nan")

    df = df.sort_values(["order_date", "order_id"], kind="stable").reset_index(drop=True)
    # Flag the first line of each order. Summing this flag over any slice gives distinct
    # orders, which keeps order counts additive for Excel SUMIFS formulas.
    df["order_count"] = (~df["order_id"].duplicated()).astype(int)
    # Weeks run Monday to Sunday and are labelled by their Sunday.
    df["week_ending"] = df["order_date"] + pd.to_timedelta((6 - df["order_date"].dt.weekday) % 7, unit="D")
    df["month"] = df["order_date"].dt.to_period("M").dt.to_timestamp()

    if rows_dropped:
        log.warning("Dropped %d of %d rows with a missing date, sales value, order ID, region or category",
                    rows_dropped, rows_read)
    log.info("Loaded %d rows (%d orders) from %s, covering %s to %s",
             len(df), int(df["order_count"].sum()), path.name,
             f"{df['order_date'].min():%Y-%m-%d}", f"{df['order_date'].max():%Y-%m-%d}")
    return LoadResult(df[CLEAN_COLUMNS], path, rows_read, rows_dropped)


def _read_raw(path: Path) -> pd.DataFrame:
    suffix = path.suffix.lower()
    if suffix == ".csv":
        # Superstore-style CSVs are often Windows-1252 encoded, so fall back gracefully.
        for encoding in ("utf-8-sig", "cp1252"):
            try:
                return pd.read_csv(path, encoding=encoding)
            except UnicodeDecodeError:
                continue
        return pd.read_csv(path, encoding="latin-1")
    if suffix in {".xlsx", ".xls"}:
        workbook = pd.ExcelFile(path)
        sheet = "Orders" if "Orders" in workbook.sheet_names else workbook.sheet_names[0]
        return workbook.parse(sheet)
    raise ValueError(f"Unsupported file type '{path.suffix}'. Use a .csv, .xlsx or .xls file.")


def _match_columns(columns) -> dict[str, str]:
    """Map raw column names to standard field names."""
    lookup = {str(c).strip().lower(): c for c in columns}
    mapping: dict[str, str] = {}
    missing = []
    for field, aliases in REQUIRED_COLUMNS.items():
        source = next((lookup[a] for a in aliases if a in lookup), None)
        if source is None:
            missing.append(field)
        else:
            mapping[source] = field
    if missing:
        raise ValueError(
            f"Missing required column(s): {', '.join(missing)}. "
            f"Columns found: {', '.join(map(str, columns))}"
        )
    for field, aliases in OPTIONAL_COLUMNS.items():
        source = next((lookup[a] for a in aliases if a in lookup), None)
        if source is not None:
            mapping[source] = field
    return mapping

"""Weekly refresh: python -m sales_pack --input data/raw/superstore.csv"""

from __future__ import annotations

import argparse
import logging
import shutil
import sys
from pathlib import Path

import pandas as pd

from .config import DEFAULT_GROWTH, TREND_WEEKS
from .excel_pack import RunInfo, write_pack
from .insights import build_callouts, money, pct
from .load import load_sales
from .metrics import category_mix, regional_table, subcategory_drivers, weekly_trend
from .periods import build_periods, default_as_of
from .quotas import build_quotas
from .tableau_feed import write_tableau_feed

log = logging.getLogger("sales_pack")


def parse_args(argv=None) -> argparse.Namespace:
    parser = argparse.ArgumentParser(
        prog="python -m sales_pack",
        description="Build the weekly regional sales pack (Excel) and the Tableau feed (CSV).",
    )
    parser.add_argument("--input", "-i", required=True,
                        help="Raw sales file (.csv, .xlsx or .xls), e.g. data/raw/superstore.csv")
    parser.add_argument("--as-of", help="Week-ending date for the review, YYYY-MM-DD. "
                                        "Default: the last complete week (Sunday) in the data.")
    parser.add_argument("--growth", type=float, default=DEFAULT_GROWTH,
                        help="Quota growth vs the same month last year, e.g. 0.10 for 10%% (default 0.10).")
    parser.add_argument("--quotas", help="Optional CSV with columns region,month,quota that overrides derived quotas.")
    parser.add_argument("--out", default="output", help="Output folder (default: output).")
    return parser.parse_args(argv)


def setup_logging(log_dir: Path) -> None:
    log_dir.mkdir(parents=True, exist_ok=True)
    logging.basicConfig(
        level=logging.INFO,
        format="%(asctime)s %(levelname)-7s %(message)s",
        handlers=[logging.StreamHandler(sys.stdout),
                  logging.FileHandler(log_dir / "refresh.log", encoding="utf-8")],
    )


def run(args: argparse.Namespace) -> Path:
    loaded = load_sales(args.input)
    sales = loaded.data
    first, last = sales["order_date"].min(), sales["order_date"].max()

    as_of = pd.Timestamp(args.as_of).normalize() if args.as_of else default_as_of(last)
    if not first <= as_of <= last + pd.Timedelta(days=6):
        raise ValueError(f"--as-of {as_of:%Y-%m-%d} is outside the data ({first:%Y-%m-%d} to {last:%Y-%m-%d}).")
    if as_of.weekday() != 6:
        log.warning("%s is not a Sunday; the week window still covers the 7 days ending on it.", f"{as_of:%Y-%m-%d}")
    if as_of < first + pd.DateOffset(years=1):
        log.warning("The as-of date is within the first year of data, so quota and last-year comparisons are blank.")

    periods = build_periods(as_of)
    quotas = build_quotas(sales, args.growth, args.quotas)
    regional = {p.key: regional_table(sales, quotas, p) for p in periods}
    trend = weekly_trend(sales, as_of, TREND_WEEKS)
    qtd = next(p for p in periods if p.key == "qtd")
    mix = category_mix(sales, qtd)
    drivers = subcategory_drivers(sales, qtd)
    callouts = build_callouts(regional, trend, mix, drivers)

    out_dir = Path(args.out)
    out_dir.mkdir(parents=True, exist_ok=True)
    quota_method = (f"Override file {Path(args.quotas).name}, with derived quotas where it has no value"
                    if args.quotas else
                    f"Derived: same month last year × (1 + {args.growth:.0%})")
    info = RunInfo(
        source=str(loaded.source), rows_loaded=len(sales), rows_dropped=loaded.rows_dropped,
        first_date=first, last_date=last, as_of=as_of, growth=args.growth,
        quota_method=quota_method, trend_weeks=TREND_WEEKS,
    )
    pack = write_pack(
        out_dir / f"Sales_Pack_{as_of:%Y-%m-%d}.xlsx", sales, quotas, periods,
        regions=sorted(sales["region"].unique()), categories=list(mix.index),
        subcategories=drivers[["sub_category", "category"]], callouts=callouts, info=info,
    )
    shutil.copyfile(pack, out_dir / "Sales_Pack_latest.xlsx")
    feed = write_tableau_feed(out_dir / "tableau", sales, quotas)

    total = regional["qtd"].loc["Total"]
    log.info("Week ending %s | QTD revenue %s | attainment %s | growth vs PY %s",
             f"{as_of:%Y-%m-%d}", money(total["revenue"]), pct(total["attainment"]),
             pct(total["growth_vs_py"], signed=True))
    log.info("Excel pack: %s (copied to Sales_Pack_latest.xlsx)", pack)
    log.info("Tableau feed: %s", ", ".join(str(p) for p in feed))
    print("\nKey callouts:")
    for line in callouts:
        print(f"  - {line}")
    return pack


def main(argv=None) -> int:
    args = parse_args(argv)
    setup_logging(Path("logs"))
    try:
        run(args)
        return 0
    except PermissionError as exc:
        log.error("Couldn't write %s. Close the file if it's open in Excel, then run again.", exc.filename)
    except (FileNotFoundError, ValueError) as exc:
        log.error("%s", exc)
    return 1


if __name__ == "__main__":
    sys.exit(main())

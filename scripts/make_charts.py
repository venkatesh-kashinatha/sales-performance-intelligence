"""Build the README charts in docs/ from a sales file, using the same metrics as the pack.

    python scripts/make_charts.py --input data/raw/superstore.csv [--as-of YYYY-MM-DD]
"""
import argparse
import sys
from pathlib import Path

import matplotlib
matplotlib.use("Agg")
import matplotlib.pyplot as plt
import matplotlib.ticker as mtick
import pandas as pd

ROOT = Path(__file__).resolve().parents[1]
sys.path.insert(0, str(ROOT))
from sales_pack.config import DEFAULT_GROWTH, TREND_WEEKS  # noqa: E402
from sales_pack.load import load_sales  # noqa: E402
from sales_pack.metrics import category_mix, regional_table, subcategory_drivers, weekly_trend  # noqa: E402
from sales_pack.periods import build_periods, default_as_of  # noqa: E402
from sales_pack.quotas import build_quotas  # noqa: E402

DOCS = ROOT / "docs"
GOOD, AMBER, BAD, INK = "#2E7D5B", "#E0A100", "#C0392B", "#1F2933"
plt.rcParams.update({"font.size": 10, "axes.spines.top": False, "axes.spines.right": False,
                     "axes.titleweight": "bold", "axes.titlesize": 12, "figure.dpi": 130})


def main():
    ap = argparse.ArgumentParser()
    ap.add_argument("--input", required=True)
    ap.add_argument("--as-of")
    ap.add_argument("--growth", type=float, default=DEFAULT_GROWTH)
    a = ap.parse_args()
    sales = load_sales(a.input).data
    as_of = pd.Timestamp(a.as_of) if a.as_of else default_as_of(sales["order_date"].max())
    quotas = build_quotas(sales, a.growth)
    periods = {p.key: p for p in build_periods(as_of)}
    DOCS.mkdir(exist_ok=True)
    qtd = periods["qtd"]

    # 1. Attainment by region, quarter to date
    t = regional_table(sales, quotas, qtd).drop(index="Total").sort_values("attainment")
    fig, ax = plt.subplots(figsize=(7, 3.8))
    colors = [GOOD if x >= 1 else (AMBER if x >= 0.9 else BAD) for x in t["attainment"]]
    ax.barh(t.index, t["attainment"], color=colors)
    ax.axvline(1, color=INK, lw=1, ls="--")
    for y, (att, gap) in enumerate(zip(t["attainment"], t["gap_to_quota"])):
        ax.text(att + 0.02, y, f"{att:.0%}  ({'+' if gap >= 0 else '-'}${abs(gap)/1000:,.1f}K vs pace)",
                va="center", fontsize=8)
    ax.xaxis.set_major_formatter(mtick.PercentFormatter(1.0))
    ax.set_xlim(0, t["attainment"].max() + 0.6)
    ax.set_title(f"Quota attainment by region, quarter to date ({as_of:%b %d, %Y})")
    fig.tight_layout(); fig.savefig(DOCS / "attainment_by_region.png"); plt.close(fig)

    # 2. Weekly revenue vs same week last year
    w = weekly_trend(sales, as_of, TREND_WEEKS)
    fig, ax = plt.subplots(figsize=(8, 3.8))
    ax.plot(w["week_ending"], w["total"] / 1000, marker="o", ms=3, color=INK, label="This year")
    ax.plot(w["week_ending"], w["py_total"] / 1000, marker="o", ms=3, color="#8A8F98", ls="--", label="Same week last year")
    ax.set_ylabel("Revenue ($K)"); ax.set_title(f"Weekly revenue, last {TREND_WEEKS} weeks")
    ax.legend(frameon=False); fig.autofmt_xdate()
    fig.tight_layout(); fig.savefig(DOCS / "weekly_trend.png"); plt.close(fig)

    # 3. Sub-category drivers vs last year, quarter to date
    d = subcategory_drivers(sales, qtd)
    d = pd.concat([d.nlargest(5, "change"), d.nsmallest(5, "change")]).drop_duplicates().sort_values("change")
    fig, ax = plt.subplots(figsize=(7, 4.2))
    ax.barh(d["sub_category"], d["change"] / 1000, color=[GOOD if c >= 0 else BAD for c in d["change"]])
    ax.axvline(0, color=INK, lw=0.8)
    ax.set_xlabel("Change in revenue vs same dates last year ($K)")
    ax.set_title("Biggest sub-category lifts and drags, quarter to date")
    fig.tight_layout(); fig.savefig(DOCS / "subcategory_drivers.png"); plt.close(fig)

    # 4. Category mix vs last year
    m = category_mix(sales, qtd)
    fig, ax = plt.subplots(figsize=(6, 3.6))
    x = range(len(m)); wd = 0.38
    ax.bar([i - wd / 2 for i in x], m["py_share"], wd, color="#8A8F98", label="Last year")
    ax.bar([i + wd / 2 for i in x], m["share"], wd, color="#3366AA", label="This year")
    ax.set_xticks(list(x)); ax.set_xticklabels(m.index)
    ax.yaxis.set_major_formatter(mtick.PercentFormatter(1.0, decimals=0))
    ax.set_title("Sales mix by category, quarter to date"); ax.legend(frameon=False)
    fig.tight_layout(); fig.savefig(DOCS / "category_mix.png"); plt.close(fig)
    print("Charts written to", DOCS)


if __name__ == "__main__":
    main()

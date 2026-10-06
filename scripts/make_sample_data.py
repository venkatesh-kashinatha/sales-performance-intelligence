"""Generate a Superstore-style sample file so the pipeline can run before you download real data.

Use it to check your setup. For your portfolio, build the pack from the real Superstore data.

    python scripts/make_sample_data.py
    python scripts/make_sample_data.py --end 2025-12-28 --years 4 --out data/raw/sample_superstore.csv
"""

from __future__ import annotations

import argparse
from datetime import date, timedelta
from pathlib import Path

import numpy as np
import pandas as pd

# region: (states, share of orders in year one, yearly growth)
REGIONS = {
    "West": (["California", "Washington", "Colorado", "Arizona", "Oregon"], 0.32, 0.15),
    "East": (["New York", "Pennsylvania", "Ohio", "Massachusetts", "New Jersey"], 0.29, 0.10),
    "Central": (["Texas", "Illinois", "Michigan", "Indiana", "Minnesota"], 0.23, 0.03),
    "South": (["Florida", "Georgia", "North Carolina", "Virginia", "Tennessee"], 0.16, 0.08),
}
# sub-category: (category, typical unit price, popularity, margin before discounts)
SUBCATEGORIES = {
    "Bookcases": ("Furniture", 200, 2.0, 0.05), "Chairs": ("Furniture", 180, 5.0, 0.08),
    "Furnishings": ("Furniture", 40, 7.5, 0.14), "Tables": ("Furniture", 300, 2.5, -0.02),
    "Appliances": ("Office Supplies", 110, 3.5, 0.16), "Art": ("Office Supplies", 11, 6.0, 0.24),
    "Binders": ("Office Supplies", 22, 12.0, 0.15), "Envelopes": ("Office Supplies", 20, 2.0, 0.40),
    "Fasteners": ("Office Supplies", 6, 2.0, 0.30), "Labels": ("Office Supplies", 7, 3.0, 0.42),
    "Paper": ("Office Supplies", 16, 11.0, 0.42), "Storage": ("Office Supplies", 90, 6.5, 0.10),
    "Supplies": ("Office Supplies", 28, 2.0, 0.02), "Accessories": ("Technology", 65, 6.0, 0.24),
    "Copiers": ("Technology", 600, 0.5, 0.35), "Machines": ("Technology", 300, 1.0, 0.05),
    "Phones": ("Technology", 110, 7.0, 0.14),
}
CATEGORY_TREND = {"Furniture": -0.04, "Office Supplies": 0.0, "Technology": 0.09}  # yearly mix drift
MONTH_SEASONALITY = [0.70, 0.55, 1.00, 0.85, 0.90, 0.85, 0.80, 0.90, 1.40, 1.00, 1.45, 1.50]
SEGMENTS, SEGMENT_P = ["Consumer", "Corporate", "Home Office"], [0.52, 0.30, 0.18]
SHIP_MODES = {"Standard Class": (0.60, 5), "Second Class": (0.19, 3), "First Class": (0.16, 2), "Same Day": (0.05, 0)}


def last_sunday(today: date) -> date:
    return today - timedelta(days=(today.weekday() + 1) % 7)


def generate(end: date, years: int, seed: int) -> pd.DataFrame:
    rng = np.random.default_rng(seed)
    start = date(end.year - years, end.month, min(end.day, 28)) + timedelta(days=1)
    customers = [f"{a}{b}-{rng.integers(10000, 99999)}" for a in "ABCDEFGHJKLMNPRSTVW" for b in "ABCDEHKMRS"][:800]
    sub_names = list(SUBCATEGORIES)
    ship_names = list(SHIP_MODES)
    ship_p = [SHIP_MODES[m][0] for m in ship_names]

    rows, order_no, day = [], 100000, start
    while day <= end:
        t = (day - start).days / 365.25  # years since start
        weekday_factor = 0.8 if day.weekday() >= 5 else 1.05
        expected = 3.2 * MONTH_SEASONALITY[day.month - 1] * weekday_factor * (1.09 ** t)
        region_w = np.array([share * (1 + g) ** t for _, share, g in REGIONS.values()])
        region_w /= region_w.sum()
        sub_w = np.array([pop * (1 + CATEGORY_TREND[cat]) ** t for cat, _, pop, _ in SUBCATEGORIES.values()])
        sub_w /= sub_w.sum()

        for _ in range(rng.poisson(expected)):
            order_no += 1
            region = rng.choice(list(REGIONS), p=region_w)
            state = rng.choice(REGIONS[region][0])
            prefix = "CA" if rng.random() < 0.8 else "US"
            order_id = f"{prefix}-{day.year}-{order_no}"
            customer = customers[rng.integers(len(customers))]
            segment = rng.choice(SEGMENTS, p=SEGMENT_P)
            mode = rng.choice(ship_names, p=ship_p)
            ship = day + timedelta(days=int(SHIP_MODES[mode][1] + rng.integers(0, 2)))
            for _ in range(1 + min(rng.poisson(0.9), 5)):
                sub = rng.choice(sub_names, p=sub_w)
                category, price, _, margin = SUBCATEGORIES[sub]
                qty = int(1 + min(rng.poisson(2.6), 13))
                unit = price * rng.lognormal(0, 0.45)
                high_discount = region == "Central"  # Central discounts harder, as in the real data
                discount = rng.choice([0, 0.1, 0.2, 0.3, 0.4, 0.5],
                                      p=[0.42, 0.10, 0.28, 0.08, 0.07, 0.05] if high_discount
                                      else [0.58, 0.10, 0.24, 0.04, 0.02, 0.02])
                sales = round(unit * qty * (1 - discount), 2)
                profit = round(sales * margin - unit * qty * discount * 0.6, 2)
                rows.append({
                    "Order ID": order_id, "Order Date": f"{day.month}/{day.day}/{day.year}",
                    "Ship Date": f"{ship.month}/{ship.day}/{ship.year}", "Ship Mode": mode,
                    "Customer ID": customer, "Segment": segment, "Country": "United States",
                    "State": state, "Region": region, "Category": category, "Sub-Category": sub,
                    "Sales": sales, "Quantity": qty, "Discount": discount, "Profit": profit,
                })
        day += timedelta(days=1)

    df = pd.DataFrame(rows)
    df.insert(0, "Row ID", range(1, len(df) + 1))
    return df


def main() -> None:
    parser = argparse.ArgumentParser(description=__doc__, formatter_class=argparse.RawDescriptionHelpFormatter)
    parser.add_argument("--out", default="data/raw/sample_superstore.csv")
    parser.add_argument("--end", help="Last order date, YYYY-MM-DD (default: the most recent Sunday)")
    parser.add_argument("--years", type=int, default=4)
    parser.add_argument("--seed", type=int, default=7)
    args = parser.parse_args()

    end = date.fromisoformat(args.end) if args.end else last_sunday(date.today())
    df = generate(end, args.years, args.seed)
    out = Path(args.out)
    out.parent.mkdir(parents=True, exist_ok=True)
    df.to_csv(out, index=False)
    print(f"Wrote {len(df):,} order lines ({df['Order ID'].nunique():,} orders) to {out}")


if __name__ == "__main__":
    main()

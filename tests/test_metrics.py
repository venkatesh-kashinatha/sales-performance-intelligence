import pandas as pd
import pytest

from sales_pack.load import load_sales
from sales_pack.metrics import quota_to_date, regional_table, weekly_trend
from sales_pack.periods import build_periods, default_as_of, same_day_last_year
from sales_pack.quotas import build_quotas


def make_sales(rows):
    df = pd.DataFrame(rows, columns=["order_id", "order_date", "region", "category", "sub_category", "sales"])
    df["order_date"] = pd.to_datetime(df["order_date"])
    df["order_count"] = (~df["order_id"].duplicated()).astype(int)
    df["month"] = df["order_date"].dt.to_period("M").dt.to_timestamp()
    return df


SALES = make_sales([
    ("A1", "2022-10-03", "West", "Technology", "Phones", 100.0),
    ("A2", "2023-10-02", "West", "Technology", "Phones", 150.0),
    ("A2", "2023-10-02", "West", "Furniture", "Chairs", 50.0),   # second line of the same order
    ("B1", "2022-10-04", "East", "Furniture", "Tables", 80.0),
    ("B2", "2023-10-03", "East", "Furniture", "Tables", 60.0),
])


def test_default_as_of_is_last_complete_week():
    assert default_as_of("2017-12-30") == pd.Timestamp("2017-12-24")  # Saturday -> previous Sunday
    assert default_as_of("2017-12-31") == pd.Timestamp("2017-12-31")  # Sunday stays


def test_period_windows():
    p = {x.key: x for x in build_periods("2023-11-26")}
    assert p["week"].start == pd.Timestamp("2023-11-20")
    assert p["mtd"].start == pd.Timestamp("2023-11-01")
    assert p["qtd"].start == pd.Timestamp("2023-10-01")
    assert p["ytd"].start == pd.Timestamp("2023-01-01")
    assert p["week"].py_start == pd.Timestamp("2022-11-21")  # 52 weeks back: Monday to Monday
    assert p["qtd"].py_start == pd.Timestamp("2022-10-01")   # calendar dates for period-to-date


def test_leap_day_matches_excel_edate():
    assert same_day_last_year("2024-02-29") == pd.Timestamp("2023-02-28")


def test_quota_is_prorated_by_days():
    quotas = pd.DataFrame({"region": ["West", "West"],
                           "month": pd.to_datetime(["2023-10-01", "2023-11-01"]),
                           "quota": [3100.0, 3000.0]})
    q = quota_to_date(quotas, pd.Timestamp("2023-10-01"), pd.Timestamp("2023-11-10"))
    assert q["West"] == pytest.approx(3100 + 3000 * 10 / 30)  # all of October + 10 of November's 30 days


def test_derived_quota_is_last_year_plus_growth():
    quotas = build_quotas(SALES, growth=0.10)
    oct_west = quotas[(quotas["region"] == "West") & (quotas["month"] == "2023-10-01")]
    assert oct_west["quota"].item() == pytest.approx(110.0)


def test_regional_table_counts_orders_once_and_compares_to_last_year():
    quotas = build_quotas(SALES, growth=0.10)
    week = next(p for p in build_periods("2023-10-08") if p.key == "week")
    t = regional_table(SALES, quotas, week)
    assert t.loc["West", "orders"] == 1                      # two lines, one order
    assert t.loc["West", "revenue"] == pytest.approx(200.0)
    assert t.loc["West", "aov"] == pytest.approx(200.0)
    assert t.loc["West", "quota_to_date"] == pytest.approx(110.0 * 7 / 31)
    assert t.loc["West", "growth_vs_py"] == pytest.approx(1.0)  # 200 vs 100
    assert t.loc["East", "growth_vs_py"] == pytest.approx(-0.25)  # 60 vs 80
    assert t.loc["Total", "revenue"] == pytest.approx(260.0)


def test_weekly_trend_uses_52_week_offset():
    trend = weekly_trend(SALES, pd.Timestamp("2023-10-08"), weeks=1)
    assert trend.loc[0, "total"] == pytest.approx(260.0)
    assert trend.loc[0, "py_total"] == pytest.approx(180.0)  # Oct 3-9, 2022 holds A1 and B1


def test_loader_handles_superstore_csv(tmp_path):
    csv = tmp_path / "superstore.csv"
    csv.write_bytes(
        "Row ID,Order ID,Order Date,Region,Category,Sub-Category,Product Name,Sales\n"
        "1,CA-1,11/8/2016,South,Furniture,Chairs,Caf\xe9 chair,261.96\n"
        "2,CA-1,11/8/2016,South,Furniture,Chairs,Desk chair,731.94\n"
        "3,CA-2,,West,Technology,Phones,Phone,10\n".encode("cp1252")
    )
    result = load_sales(csv)
    df = result.data
    assert result.rows_dropped == 1                          # the row without a date
    assert df["order_count"].sum() == 1
    assert df.loc[0, "week_ending"] == pd.Timestamp("2016-11-13")  # Tue Nov 8 -> Sun Nov 13
    assert df["sales"].sum() == pytest.approx(993.90)

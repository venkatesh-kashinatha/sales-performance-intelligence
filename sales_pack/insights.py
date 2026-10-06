"""Plain-English callouts for the weekly review, written from the computed metrics."""

from __future__ import annotations

import math

import pandas as pd


def money(x: float) -> str:
    sign = "-" if x < 0 else ""
    x = abs(x)
    if x >= 1_000_000:
        return f"{sign}${x / 1_000_000:.2f}M"
    if x >= 1_000:
        return f"{sign}${x / 1_000:.1f}K"
    return f"{sign}${x:,.0f}"


def pct(x, signed: bool = False) -> str:
    if x is None or (isinstance(x, float) and math.isnan(x)):
        return "n/a"
    return f"{x:+.1%}" if signed else f"{x:.0%}"


def _has(x) -> bool:
    return x is not None and not (isinstance(x, float) and math.isnan(x))


def build_callouts(regional: dict[str, pd.DataFrame], trend: pd.DataFrame,
                   mix: pd.DataFrame, drivers: pd.DataFrame) -> list[str]:
    """Return 4-7 short sentences a sales lead could read out in the weekly review."""
    qtd = regional["qtd"]
    total = qtd.loc["Total"]
    regions = qtd.drop(index="Total")
    has_quota = total["quota_to_date"] > 0
    lines: list[str] = []

    # 1. Headline: where the quarter stands
    if has_quota:
        gap = total["gap_to_quota"]
        lines.append(
            f"Quarter to date, revenue is {money(total['revenue'])}: {pct(total['attainment'])} of quota to date "
            f"({money(abs(gap))} {'ahead of' if gap >= 0 else 'behind'} pace) and "
            f"{pct(total['growth_vs_py'], signed=True)} vs the same dates last year."
        )
    else:
        lines.append(
            f"Quarter to date, revenue is {money(total['revenue'])} "
            f"({pct(total['growth_vs_py'], signed=True)} vs last year). No quota is available for this period."
        )

    # 2. Best and worst region against quota
    if has_quota and len(regions) > 1:
        best, worst = regions["attainment"].idxmax(), regions["attainment"].idxmin()
        worst_gap = regions.at[worst, "gap_to_quota"]
        lines.append(
            f"{best} leads at {pct(regions.at[best, 'attainment'])} of quota to date; {worst} trails at "
            f"{pct(regions.at[worst, 'attainment'])} ({money(abs(worst_gap))} "
            f"{'behind' if worst_gap < 0 else 'ahead of'} pace)."
        )

    # 3. Last week vs the same week last year
    last = trend.iloc[-1]
    lines.append(
        f"Week ending {last['week_ending']:%b %d}: {money(last['total'])} in revenue, "
        f"{pct(last['growth_vs_py'], signed=True)} vs the same week last year."
    )

    # 4. Biggest category mix shift
    if mix["mix_shift_pts"].notna().any():
        cat = mix["mix_shift_pts"].abs().idxmax()
        pts = mix.at[cat, "mix_shift_pts"]
        lines.append(
            f"Biggest mix shift: {cat} is {abs(pts):.1f} pts {'higher' if pts >= 0 else 'lower'} as a share "
            f"of revenue than last year ({pct(mix.at[cat, 'share'])} of revenue now)."
        )

    # 5. Sub-categories driving the change
    if len(drivers):
        drag = drivers.loc[drivers["change"].idxmin()]
        lift = drivers.loc[drivers["change"].idxmax()]
        if drag["change"] < 0:
            lines.append(f"Biggest drag: {drag['sub_category']} ({drag['category']}), "
                         f"down {money(abs(drag['change']))} vs last year quarter to date.")
        if lift["change"] > 0:
            lines.append(f"Biggest lift: {lift['sub_category']} ({lift['category']}), "
                         f"up {money(lift['change'])} vs last year quarter to date.")

    # 6. Average order value
    if _has(total["aov"]) and _has(total["py_aov"]) and total["py_aov"] > 0:
        change = total["aov"] / total["py_aov"] - 1
        lines.append(f"Average order value is {money(total['aov'])} quarter to date vs "
                     f"{money(total['py_aov'])} last year ({pct(change, signed=True)}).")
    return lines

"""Run the price-elasticity analysis and write everything to one Excel file.

Reads the validated CSV produced by validate_data.py, estimates a price
elasticity per product from its historical price/volume pairs, and writes a
brand-new workbook (no template) with these sheets:
  - Input Data              the validated transactions the analysis ran on
  - Elasticity by Product   regression results, current margin, recommendation per product
  - Safe to Raise           inelastic products (|elasticity| < 1) that can carry a price increase
  - Price Sensitive         elastic products (|elasticity| > 1) where discounts can pay off
  - Promo Scenarios         markdown depth vs. expected lift vs. break-even lift, per product
  - Price Change Simulation contribution impact of -20% .. +20% price moves, per product
  - Method Notes            plain-language explanation of every number in the workbook

Method:
  - Deseasonalise:     units are divided by a calendar-month index built from total
                       units across all products, so holiday peaks aren't mistaken for
                       a response to the promotions that tend to run in the same weeks.
  - Log-log regression: ln(units) = a + b * ln(price), fitted per product by ordinary
                       least squares. The slope b is the price elasticity of demand
                       (% change in units for a 1% change in price).
  - Contribution:      weekly contribution = units * (price - unit cost). Every scenario is
                       judged on contribution, never on revenue.
  - Break-even lift:   the volume increase needed for a discount d to leave contribution
                       unchanged: (P - c) / (P * (1 - d) - c) - 1. A promo only pays if the
                       expected lift (1 - d) ** b - 1 exceeds this.
  - Optimal margin:    for elastic products the contribution-maximising margin is
                       (P - c) / P = -1 / b (the Lerner rule). Inelastic products have no
                       interior optimum: raising price always raises contribution, so
                       they're flagged "safe to raise" with a bounded test increase.
"""

import argparse
import csv
import math
import os
from collections import defaultdict
from datetime import date
from statistics import mean

from openpyxl import Workbook
from openpyxl.styles import Alignment, Font, PatternFill
from openpyxl.utils import get_column_letter

from pricing_common import (
    CSV_FIELDS,
    MIN_ABS_T_STAT,
    MIN_OBSERVATIONS,
    MIN_PRICE_POINTS,
    OUTPUT_WORKBOOK_PATH,
    PRICE_CHANGE_STEPS,
    PROMO_DEPTHS,
    TRUTH_DATA_PATH,
    UNIT_ELASTICITY,
    VALIDATED_DATA_PATH,
    Transaction,
)

HEADER_FILL = PatternFill(start_color="1F4E78", end_color="1F4E78", fill_type="solid")
HEADER_FONT = Font(color="FFFFFF", bold=True)
GREEN_FILL = PatternFill(start_color="C6EFCE", end_color="C6EFCE", fill_type="solid")
ORANGE_FILL = PatternFill(start_color="FFEB9C", end_color="FFEB9C", fill_type="solid")
RED_FILL = PatternFill(start_color="FFC7CE", end_color="FFC7CE", fill_type="solid")
GREY_FILL = PatternFill(start_color="E7E6E6", end_color="E7E6E6", fill_type="solid")

SAFE_TO_RAISE = "Safe to raise"
PRICE_SENSITIVE = "Price-sensitive"
INCONCLUSIVE = "Inconclusive"
INSUFFICIENT = "Insufficient data"

CLASS_FILLS = {
    SAFE_TO_RAISE: GREEN_FILL,
    PRICE_SENSITIVE: ORANGE_FILL,
    INCONCLUSIVE: GREY_FILL,
    INSUFFICIENT: GREY_FILL,
}

ELASTICITY_FIELDS = [
    "sku",
    "product_name",
    "category",
    "classification",
    "elasticity",
    "std_error",
    "t_stat",
    "r_squared",
    "n_weeks",
    "n_price_points",
    "min_price",
    "max_price",
    "current_price",
    "unit_cost",
    "unit_margin",
    "margin_pct",
    "optimal_margin_pct",
    "weekly_units_at_current",
    "weekly_contribution_at_current",
    "best_price_change_pct",
    "best_price",
    "contribution_change_pct",
    "max_profitable_markdown_pct",
    "recommendation",
]

PROMO_FIELDS = [
    "sku",
    "product_name",
    "classification",
    "elasticity",
    "current_price",
    "unit_cost",
    "markdown_pct",
    "promo_price",
    "promo_unit_margin",
    "break_even_lift_pct",
    "expected_lift_pct",
    "lift_gap_pct",
    "baseline_weekly_units",
    "promo_weekly_units",
    "baseline_weekly_contribution",
    "promo_weekly_contribution",
    "contribution_change_pct",
    "verdict",
]

SIMULATION_FIELDS = [
    "sku",
    "product_name",
    "classification",
    "elasticity",
    "current_price",
    "price_change_pct",
    "new_price",
    "expected_units_change_pct",
    "break_even_units_change_pct",
    "expected_weekly_units",
    "expected_weekly_revenue",
    "expected_weekly_contribution",
    "contribution_change_pct",
    "verdict",
]


# ---------------------------------------------------------------------------
# Loading and preparing observations
# ---------------------------------------------------------------------------


def read_transactions(path: str) -> list[Transaction]:
    with open(path, newline="", encoding="utf-8") as f:
        return [Transaction(**row) for row in csv.DictReader(f)]


def read_truth(path: str | None) -> dict[str, float]:
    if not path or not os.path.exists(path):
        return {}
    with open(path, newline="", encoding="utf-8") as f:
        return {row["sku"]: float(row["true_elasticity"]) for row in csv.DictReader(f)}


def build_weekly_observations(txns: list[Transaction]) -> dict[str, list[dict]]:
    """Collapse channels: one (week, price) -> units observation per product."""
    grouped: dict[tuple[str, date, float], dict] = {}
    for t in txns:
        key = (t.sku, t.week_start_date, round(t.unit_price, 2))
        obs = grouped.setdefault(
            key,
            {
                "sku": t.sku,
                "product_name": t.product_name,
                "category": t.category,
                "week_start_date": t.week_start_date,
                "unit_price": round(t.unit_price, 2),
                "unit_cost": t.unit_cost,
                "on_promo": t.on_promo,
                "units_sold": 0,
            },
        )
        obs["units_sold"] += t.units_sold
    by_sku: dict[str, list[dict]] = defaultdict(list)
    for obs in grouped.values():
        by_sku[obs["sku"]].append(obs)
    for obs_list in by_sku.values():
        obs_list.sort(key=lambda o: o["week_start_date"])
    return by_sku


def seasonal_index_by_month(by_sku: dict[str, list[dict]]) -> dict[int, float]:
    """Average weekly units per calendar month across all products, relative to the overall mean."""
    month_units: dict[int, list[int]] = defaultdict(list)
    for obs_list in by_sku.values():
        for obs in obs_list:
            month_units[obs["week_start_date"].month].append(obs["units_sold"])
    month_means = {m: mean(u) for m, u in month_units.items() if u}
    overall = mean(month_means.values()) if month_means else 1.0
    return {m: (v / overall if overall > 0 else 1.0) for m, v in month_means.items()}


# ---------------------------------------------------------------------------
# Regression
# ---------------------------------------------------------------------------


def ols(xs: list[float], ys: list[float]) -> dict | None:
    """Simple ordinary-least-squares fit y = a + b x, in pure Python."""
    n = len(xs)
    if n < 3:
        return None
    mx, my = mean(xs), mean(ys)
    sxx = sum((x - mx) ** 2 for x in xs)
    if sxx == 0:
        return None
    sxy = sum((x - mx) * (y - my) for x, y in zip(xs, ys))
    b = sxy / sxx
    a = my - b * mx
    sse = sum((y - (a + b * x)) ** 2 for x, y in zip(xs, ys))
    syy = sum((y - my) ** 2 for y in ys)
    r_squared = 1 - sse / syy if syy > 0 else 0.0
    se_b = math.sqrt(sse / (n - 2) / sxx) if n > 2 else float("inf")
    t_stat = b / se_b if se_b > 0 else float("inf")
    return {"intercept": a, "slope": b, "std_error": se_b, "t_stat": t_stat, "r_squared": r_squared, "n": n}


def fit_elasticity(obs_list: list[dict], season_index: dict[int, float]) -> dict:
    usable = [o for o in obs_list if o["units_sold"] > 0]
    xs = [math.log(o["unit_price"]) for o in usable]
    ys = [math.log(o["units_sold"] / season_index.get(o["week_start_date"].month, 1.0)) for o in usable]
    fit = ols(xs, ys)
    price_points = {o["unit_price"] for o in usable}
    return {
        "fit": fit,
        "n_weeks": len(usable),
        "n_price_points": len(price_points),
        "min_price": min(price_points) if price_points else None,
        "max_price": max(price_points) if price_points else None,
    }


def classify(fit: dict | None, n_weeks: int, n_price_points: int) -> str:
    if fit is None or n_weeks < MIN_OBSERVATIONS or n_price_points < MIN_PRICE_POINTS:
        return INSUFFICIENT
    if fit["slope"] >= 0 or abs(fit["t_stat"]) < MIN_ABS_T_STAT:
        return INCONCLUSIVE
    return SAFE_TO_RAISE if fit["slope"] > UNIT_ELASTICITY else PRICE_SENSITIVE


# ---------------------------------------------------------------------------
# Economics
# ---------------------------------------------------------------------------


def predicted_units(fit: dict, price: float) -> float:
    return math.exp(fit["intercept"] + fit["slope"] * math.log(price))


def contribution(units: float, price: float, cost: float) -> float:
    return units * (price - cost)


def break_even_units_change(price: float, cost: float, change: float) -> float | None:
    """Volume change (fraction) that keeps contribution flat after a price change.

    Negative `change` is a discount (needs a lift); positive is an increase
    (tolerates a loss). Returns None if the new price is at or below cost.
    """
    new_margin = price * (1 + change) - cost
    if new_margin <= 0:
        return None
    return (price - cost) / new_margin - 1


def expected_units_change(elasticity: float, change: float) -> float:
    return (1 + change) ** elasticity - 1


def pct(x: float | None, digits: int = 1) -> float | int | str:
    if x is None:
        return ""
    return int(round(x * 100)) if digits == 0 else round(x * 100, digits)


def current_price_and_cost(obs_list: list[dict]) -> tuple[float, float]:
    """Latest non-promo (list) price, falling back to the latest price of any kind."""
    non_promo = [o for o in obs_list if not o["on_promo"]]
    latest = (non_promo or obs_list)[-1]
    return latest["unit_price"], obs_list[-1]["unit_cost"]


def analyze_product(sku: str, obs_list: list[dict], season_index: dict[int, float]) -> dict:
    stats = fit_elasticity(obs_list, season_index)
    fit = stats["fit"]
    classification = classify(fit, stats["n_weeks"], stats["n_price_points"])
    price, cost = current_price_and_cost(obs_list)
    first = obs_list[0]

    result = {
        "sku": sku,
        "product_name": first["product_name"],
        "category": first["category"],
        "classification": classification,
        "elasticity": round(fit["slope"], 3) if fit else "",
        "std_error": round(fit["std_error"], 3) if fit else "",
        "t_stat": round(fit["t_stat"], 2) if fit else "",
        "r_squared": round(fit["r_squared"], 3) if fit else "",
        "n_weeks": stats["n_weeks"],
        "n_price_points": stats["n_price_points"],
        "min_price": stats["min_price"],
        "max_price": stats["max_price"],
        "current_price": price,
        "unit_cost": cost,
        "unit_margin": round(price - cost, 2),
        "margin_pct": pct((price - cost) / price),
        "optimal_margin_pct": "",
        "weekly_units_at_current": "",
        "weekly_contribution_at_current": "",
        "best_price_change_pct": "",
        "best_price": "",
        "contribution_change_pct": "",
        "max_profitable_markdown_pct": "",
        "recommendation": "",
        # Not written to the sheet; carried for the scenario tables.
        "_fit": fit,
    }

    if classification in (INSUFFICIENT, INCONCLUSIVE):
        result["recommendation"] = (
            "Too few distinct price points or weeks to estimate elasticity - run a controlled price test."
            if classification == INSUFFICIENT
            else "Elasticity not statistically distinguishable from zero - keep price, gather more variation."
        )
        return result

    b = fit["slope"]
    base_units = predicted_units(fit, price)
    base_contrib = contribution(base_units, price, cost)
    result["weekly_units_at_current"] = round(base_units, 1)
    result["weekly_contribution_at_current"] = round(base_contrib, 2)
    if b < UNIT_ELASTICITY:
        result["optimal_margin_pct"] = pct(-1 / b)

    # Best move within the simulated range, judged on weekly contribution.
    best_change, best_contrib = 0.0, base_contrib
    for change in PRICE_CHANGE_STEPS:
        new_price = price * (1 + change)
        c = contribution(predicted_units(fit, new_price), new_price, cost)
        if c > best_contrib:
            best_change, best_contrib = change, c
    result["best_price_change_pct"] = pct(best_change, 0)
    result["best_price"] = round(price * (1 + best_change), 2)
    result["contribution_change_pct"] = pct(best_contrib / base_contrib - 1) if base_contrib > 0 else ""

    # Deepest markdown whose expected lift still beats the break-even lift.
    profitable = [
        d
        for d in PROMO_DEPTHS
        if (be := break_even_units_change(price, cost, -d)) is not None and expected_units_change(b, -d) > be
    ]
    result["max_profitable_markdown_pct"] = pct(max(profitable), 0) if profitable else 0

    if classification == SAFE_TO_RAISE:
        result["recommendation"] = (
            f"Demand is inelastic ({b:.2f}): every simulated increase up to +{result['best_price_change_pct']:.0f}% "
            f"adds contribution - start with a +5-10% test and re-measure. Keep this product off the promo list."
        )
    else:
        result["recommendation"] = (
            f"Demand is elastic ({b:.2f}): target a {result['optimal_margin_pct']:.0f}% margin; "
            + (
                f"markdowns up to {result['max_profitable_markdown_pct']}% add contribution."
                if profitable
                else "even a 5% markdown loses contribution at the current margin."
            )
        )
    return result


def build_promo_scenarios(results: list[dict]) -> list[dict]:
    rows = []
    for r in results:
        fit = r["_fit"]
        if r["classification"] in (INSUFFICIENT, INCONCLUSIVE):
            continue
        price, cost, b = r["current_price"], r["unit_cost"], fit["slope"]
        base_units = predicted_units(fit, price)
        base_contrib = contribution(base_units, price, cost)
        for d in PROMO_DEPTHS:
            promo_price = round(price * (1 - d), 2)
            be = break_even_units_change(price, cost, -d)
            lift = expected_units_change(b, -d)
            promo_units = base_units * (1 + lift)
            promo_contrib = contribution(promo_units, promo_price, cost)
            if be is None:
                verdict = "Below cost"
            elif lift > be:
                verdict = "Profitable"
            else:
                verdict = "Loses contribution"
            rows.append(
                {
                    "sku": r["sku"],
                    "product_name": r["product_name"],
                    "classification": r["classification"],
                    "elasticity": r["elasticity"],
                    "current_price": price,
                    "unit_cost": cost,
                    "markdown_pct": pct(d, 0),
                    "promo_price": promo_price,
                    "promo_unit_margin": round(promo_price - cost, 2),
                    "break_even_lift_pct": pct(be),
                    "expected_lift_pct": pct(lift),
                    "lift_gap_pct": pct(lift - be) if be is not None else "",
                    "baseline_weekly_units": round(base_units, 1),
                    "promo_weekly_units": round(promo_units, 1),
                    "baseline_weekly_contribution": round(base_contrib, 2),
                    "promo_weekly_contribution": round(promo_contrib, 2),
                    "contribution_change_pct": pct(promo_contrib / base_contrib - 1) if base_contrib > 0 else "",
                    "verdict": verdict,
                }
            )
    return rows


def build_price_simulation(results: list[dict]) -> list[dict]:
    rows = []
    for r in results:
        fit = r["_fit"]
        if r["classification"] in (INSUFFICIENT, INCONCLUSIVE):
            continue
        price, cost, b = r["current_price"], r["unit_cost"], fit["slope"]
        base_contrib = contribution(predicted_units(fit, price), price, cost)
        for change in PRICE_CHANGE_STEPS:
            new_price = round(price * (1 + change), 2)
            units = predicted_units(fit, new_price)
            contrib = contribution(units, new_price, cost)
            be = break_even_units_change(price, cost, change)
            rows.append(
                {
                    "sku": r["sku"],
                    "product_name": r["product_name"],
                    "classification": r["classification"],
                    "elasticity": r["elasticity"],
                    "current_price": price,
                    "price_change_pct": pct(change, 0),
                    "new_price": new_price,
                    "expected_units_change_pct": pct(expected_units_change(b, change)),
                    "break_even_units_change_pct": pct(be),
                    "expected_weekly_units": round(units, 1),
                    "expected_weekly_revenue": round(units * new_price, 2),
                    "expected_weekly_contribution": round(contrib, 2),
                    "contribution_change_pct": pct(contrib / base_contrib - 1) if base_contrib > 0 else "",
                    "verdict": "Improves contribution" if contrib > base_contrib else "Reduces contribution",
                }
            )
    return rows


# ---------------------------------------------------------------------------
# Excel output
# ---------------------------------------------------------------------------

METHOD_NOTES = [
    ("What this workbook answers", "Which products can carry a price increase, which are price-sensitive, and how deep a holiday markdown can go before it destroys contribution."),
    ("Elasticity", "Slope b of the regression ln(units) = a + b ln(price), fitted per product on its weekly price/volume history. b = -1.5 means a 1% price cut lifts units by about 1.5%."),
    ("Deseasonalising", "Weekly units are divided by a calendar-month index (total units across all products) before fitting, so the November/December peak is not credited to the promotions that tend to run then."),
    ("Std error / t-stat / R^2", "How precisely b is estimated. |t| below 2 means the data cannot distinguish the elasticity from zero, and the product is marked Inconclusive."),
    ("n_weeks / n_price_points", "Weeks with sales and distinct prices seen. Fewer than 12 weeks or 3 prices is marked Insufficient data - run a deliberate price test before acting."),
    ("Safe to raise", "b between -1 and 0 (inelastic). A price increase loses fewer units than it gains in margin, so contribution rises. Discounts on these products lose money."),
    ("Price-sensitive", "b below -1 (elastic). Volume reacts strongly to price. Optimal margin % = -1 / b (Lerner rule); if current margin is well above this, a cut can add contribution."),
    ("Contribution", "Weekly units x (price - unit cost). All verdicts use contribution, not revenue: a promo that lifts revenue can still lose money."),
    ("Break-even lift", "Volume increase needed for a discount d to leave contribution unchanged: (P - c) / (P(1 - d) - c) - 1. The promo pays only if the expected lift (1 - d)^b - 1 beats it."),
    ("Promo Scenarios sheet", "One row per product and markdown depth (5-30%). 'Profitable' means expected lift exceeds break-even lift; 'Below cost' means the promo price is at or under unit cost."),
    ("Price Change Simulation sheet", "Contribution impact of -20% to +20% price moves. The best move within this range feeds best_price on the Elasticity sheet. Do not extrapolate beyond the min/max price actually observed."),
    ("Caveats", "Constant elasticity is an approximation; competitor moves, stock-outs, and display effects are ignored. Treat outputs as a shortlist for controlled price tests, not as final prices."),
]


def autosize_columns(ws) -> None:
    for col_cells in ws.columns:
        length = max((len(str(c.value)) for c in col_cells if c.value is not None), default=10)
        ws.column_dimensions[get_column_letter(col_cells[0].column)].width = min(length + 2, 60)


def write_sheet(ws, fieldnames: list[str], rows: list[dict]) -> None:
    ws.append(fieldnames)
    for cell in ws[1]:
        cell.fill = HEADER_FILL
        cell.font = HEADER_FONT
        cell.alignment = Alignment(horizontal="center")
    for row in rows:
        ws.append([row.get(f, "") for f in fieldnames])
    ws.freeze_panes = "A2"
    if rows:
        ws.auto_filter.ref = f"A1:{get_column_letter(len(fieldnames))}{len(rows) + 1}"
    autosize_columns(ws)


def highlight_by_value(ws, fieldnames: list[str], field: str, fills: dict[str, PatternFill]) -> None:
    col = fieldnames.index(field)
    for row_cells in ws.iter_rows(min_row=2, max_row=ws.max_row):
        fill = fills.get(row_cells[col].value)
        if fill is not None:
            for cell in row_cells:
                cell.fill = fill


def write_notes_sheet(ws) -> None:
    ws.append(["Topic", "Explanation"])
    for cell in ws[1]:
        cell.fill = HEADER_FILL
        cell.font = HEADER_FONT
    for topic, text in METHOD_NOTES:
        ws.append([topic, text])
    ws.column_dimensions["A"].width = 32
    ws.column_dimensions["B"].width = 120
    for row_cells in ws.iter_rows(min_row=2, max_row=ws.max_row):
        row_cells[0].font = Font(bold=True)
        row_cells[1].alignment = Alignment(wrap_text=True, vertical="top")
        row_cells[0].alignment = Alignment(vertical="top")


def build_workbook(
    input_rows: list[dict], results: list[dict], promos: list[dict], simulation: list[dict], truth: dict[str, float]
) -> Workbook:
    wb = Workbook()

    ws_input = wb.active
    ws_input.title = "Input Data"
    write_sheet(ws_input, CSV_FIELDS, input_rows)

    fields = list(ELASTICITY_FIELDS)
    if truth:
        fields.insert(fields.index("std_error"), "true_elasticity")
        for r in results:
            r["true_elasticity"] = truth.get(r["sku"], "")

    ws_elastic = wb.create_sheet("Elasticity by Product")
    write_sheet(ws_elastic, fields, results)
    highlight_by_value(ws_elastic, fields, "classification", CLASS_FILLS)

    ws_safe = wb.create_sheet("Safe to Raise")
    safe = sorted((r for r in results if r["classification"] == SAFE_TO_RAISE), key=lambda r: r["elasticity"], reverse=True)
    write_sheet(ws_safe, fields, safe)

    ws_sensitive = wb.create_sheet("Price Sensitive")
    sensitive = sorted((r for r in results if r["classification"] == PRICE_SENSITIVE), key=lambda r: r["elasticity"])
    write_sheet(ws_sensitive, fields, sensitive)

    ws_promo = wb.create_sheet("Promo Scenarios")
    write_sheet(ws_promo, PROMO_FIELDS, promos)
    highlight_by_value(
        ws_promo, PROMO_FIELDS, "verdict", {"Profitable": GREEN_FILL, "Loses contribution": RED_FILL, "Below cost": GREY_FILL}
    )

    ws_sim = wb.create_sheet("Price Change Simulation")
    write_sheet(ws_sim, SIMULATION_FIELDS, simulation)
    highlight_by_value(
        ws_sim, SIMULATION_FIELDS, "verdict", {"Improves contribution": GREEN_FILL, "Reduces contribution": RED_FILL}
    )

    write_notes_sheet(wb.create_sheet("Method Notes"))
    return wb


# ---------------------------------------------------------------------------


def main() -> None:
    parser = argparse.ArgumentParser(description="Estimate price elasticity per product and write results to Excel.")
    parser.add_argument("--input", type=str, default=VALIDATED_DATA_PATH, help="Path to the validated CSV")
    parser.add_argument("--output", type=str, default=OUTPUT_WORKBOOK_PATH, help="Path to write the Excel workbook to")
    parser.add_argument(
        "--truth",
        type=str,
        default=TRUTH_DATA_PATH,
        help="Optional CSV of true elasticities from generate_data.py; adds a comparison column if the file exists",
    )
    args = parser.parse_args()

    txns = read_transactions(args.input)
    with open(args.input, newline="", encoding="utf-8") as f:
        input_rows = list(csv.DictReader(f))

    by_sku = build_weekly_observations(txns)
    season_index = seasonal_index_by_month(by_sku)
    results = [analyze_product(sku, obs, season_index) for sku, obs in sorted(by_sku.items())]
    results.sort(key=lambda r: (r["elasticity"] == "", r["elasticity"] if r["elasticity"] != "" else 0))

    promos = build_promo_scenarios(results)
    simulation = build_price_simulation(results)
    truth = read_truth(args.truth)

    wb = build_workbook(input_rows, results, promos, simulation, truth)
    os.makedirs(os.path.dirname(args.output) or ".", exist_ok=True)
    wb.save(args.output)

    counts = defaultdict(int)
    for r in results:
        counts[r["classification"]] += 1
    print(f"Analyzed {len(results)} products from {len(txns)} transactions -> {args.output}")
    for cls in (SAFE_TO_RAISE, PRICE_SENSITIVE, INCONCLUSIVE, INSUFFICIENT):
        print(f"  {cls:<18} {counts[cls]}")
    profitable = sum(1 for p in promos if p["verdict"] == "Profitable")
    print(f"{profitable} of {len(promos)} promo scenarios add contribution.")


if __name__ == "__main__":
    main()

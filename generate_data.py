"""Generate a fake sales-transaction dataset with Faker.

Produces a CSV of weekly product sales (price, units, cost, promo flag) over
roughly two years, in which every product is sold at several different price
points: a handful of list-price changes plus periodic promotions. Volumes are
drawn from a hidden constant-elasticity demand curve per product, so the
analysis step has a real signal to recover.

A small percentage of rows are deliberately corrupted (negative prices, blank
units, revenue that doesn't match price x units, a duplicate transaction id)
so validate_data.py has real work to do - pass --error-rate 0 for a perfectly
clean dataset.

Also writes the hidden "truth" (each product's true elasticity) to a separate
CSV so recovered elasticities can be checked against it.
"""

import argparse
import csv
import math
import os
import random
from datetime import date, timedelta

from faker import Faker

from pricing_common import CATEGORIES, CHANNELS, CSV_FIELDS, RAW_DATA_PATH, TRUTH_DATA_PATH

TRUTH_FIELDS = ["sku", "product_name", "category", "unit_cost", "base_price", "true_elasticity", "base_weekly_units"]

PRODUCT_NOUNS = ["Blend", "Pack", "Set", "Jar", "Box", "Roll", "Kit", "Bottle", "Bag", "Tin"]
PROMO_DEPTHS = [0.10, 0.15, 0.20, 0.25, 0.30]
ONLINE_SHARE = 0.30


def build_product_name(fake: Faker) -> str:
    return f"{fake.word().capitalize()} {fake.word().capitalize()} {random.choice(PRODUCT_NOUNS)}"


def generate_product(fake: Faker, index: int) -> dict:
    """A product with a hidden demand curve: units = base * (price / base_price) ** elasticity."""
    unit_cost = round(random.uniform(0.8, 45.0), 2)
    markup = random.uniform(1.35, 2.8)
    # Roughly 40% inelastic (-0.3 .. -0.95) and 60% elastic (-1.05 .. -3.2)
    # products so both the "safe to raise" and "price-sensitive" lists get
    # populated. A small gap around -1 keeps the classes distinguishable.
    if random.random() < 0.4:
        elasticity = round(random.uniform(-0.95, -0.3), 2)
    else:
        elasticity = round(random.uniform(-3.2, -1.05), 2)
    return {
        "sku": f"SKU-{index:04d}",
        "product_name": build_product_name(fake),
        "category": random.choice(CATEGORIES),
        "unit_cost": unit_cost,
        "base_price": round(unit_cost * markup, 2),
        "true_elasticity": elasticity,
        "base_weekly_units": random.randint(20, 400),
        "sells_online": random.random() < 0.6,
    }


def build_list_price_path(base_price: float, n_weeks: int) -> list[float]:
    """List price per week: starts at base_price, steps 2-4 times by +/-5-15%."""
    n_changes = random.randint(2, 4)
    change_weeks = sorted(random.sample(range(8, max(9, n_weeks - 4)), k=n_changes))
    prices: list[float] = []
    current = base_price
    for week in range(n_weeks):
        if week in change_weeks:
            current = round(current * (1 + random.choice([-1, 1]) * random.uniform(0.05, 0.15)), 2)
        prices.append(current)
    return prices


def seasonality(day: date) -> float:
    """Mild annual cycle plus a November-December holiday bump."""
    day_of_year = day.timetuple().tm_yday
    cycle = 1 + 0.15 * math.sin(2 * math.pi * (day_of_year / 365 - 0.25))
    holiday = 1.30 if day.month == 12 else 1.12 if day.month == 11 else 1.0
    return cycle * holiday


def expected_units(product: dict, price: float, week_start: date) -> float:
    ratio = price / product["base_price"]
    return product["base_weekly_units"] * ratio ** product["true_elasticity"] * seasonality(week_start)


def generate_rows(products: list[dict], n_weeks: int, start_monday: date) -> list[dict]:
    rows: list[dict] = []
    txn_no = 0
    for product in products:
        list_prices = build_list_price_path(product["base_price"], n_weeks)
        for week in range(n_weeks):
            week_start = start_monday + timedelta(weeks=week)
            promo_prob = 0.25 if week_start.month in (11, 12) else 0.10
            on_promo = random.random() < promo_prob
            price = list_prices[week]
            if on_promo:
                price = round(price * (1 - random.choice(PROMO_DEPTHS)), 2)

            mean_units = expected_units(product, price, week_start)
            channels = CHANNELS if product["sells_online"] else CHANNELS[:1]
            for channel in channels:
                if len(channels) == 1:
                    share = 1.0
                else:
                    share = ONLINE_SHARE if channel == "Online" else 1 - ONLINE_SHARE
                noise = math.exp(random.gauss(0, 0.15))
                units = max(0, round(mean_units * share * noise))
                txn_no += 1
                rows.append(
                    {
                        "transaction_id": f"TXN-{txn_no:06d}",
                        "week_start_date": week_start.isoformat(),
                        "sku": product["sku"],
                        "product_name": product["product_name"],
                        "category": product["category"],
                        "channel": channel,
                        "unit_cost": product["unit_cost"],
                        "unit_price": price,
                        "units_sold": units,
                        "revenue": round(price * units, 2),
                        "on_promo": on_promo,
                    }
                )
    return rows


def corrupt_row(row: dict) -> dict:
    """Introduce one realistic data-quality issue into a row, in place."""
    issue = random.choice(
        ["negative_price", "blank_units", "negative_cost", "revenue_mismatch", "blank_sku", "bad_date"]
    )
    if issue == "negative_price":
        row["unit_price"] = -abs(row["unit_price"])
    elif issue == "blank_units":
        row["units_sold"] = ""
    elif issue == "negative_cost":
        row["unit_cost"] = -abs(row["unit_cost"])
    elif issue == "revenue_mismatch":
        row["revenue"] = round(row["revenue"] * 1.37 + 5, 2)
    elif issue == "blank_sku":
        row["sku"] = ""
    elif issue == "bad_date":
        row["week_start_date"] = "31/13/2025"
    return row


def generate_dataset(
    n_products: int, n_weeks: int, error_rate: float, seed: int | None
) -> tuple[list[dict], list[dict]]:
    if seed is not None:
        random.seed(seed)
        Faker.seed(seed)
    fake = Faker()

    # Start on the Monday roughly n_weeks before today so the history ends "now".
    today = date.today()
    start_monday = today - timedelta(days=today.weekday()) - timedelta(weeks=n_weeks)

    products = [generate_product(fake, i) for i in range(1, n_products + 1)]
    rows = generate_rows(products, n_weeks, start_monday)

    n_errors = round(len(rows) * error_rate)
    for row in random.sample(rows, k=min(n_errors, len(rows))):
        corrupt_row(row)

    # Inject a duplicate transaction id to exercise the validator's duplicate
    # check, as long as there's more than one row to work with.
    if error_rate > 0 and len(rows) > 1:
        dup = dict(rows[0])
        dup["transaction_id"] = rows[-1]["transaction_id"]
        rows.insert(len(rows) // 2, dup)

    return rows, products


def write_csv(rows: list[dict], fieldnames: list[str], path: str) -> None:
    os.makedirs(os.path.dirname(path) or ".", exist_ok=True)
    with open(path, "w", newline="", encoding="utf-8") as f:
        writer = csv.DictWriter(f, fieldnames=fieldnames, extrasaction="ignore")
        writer.writeheader()
        writer.writerows(rows)


def main() -> None:
    parser = argparse.ArgumentParser(description="Generate fake weekly sales transactions for the pricing pipeline.")
    parser.add_argument("--products", type=int, default=40, help="Number of products to generate (default: 40)")
    parser.add_argument("--weeks", type=int, default=104, help="Weeks of history per product (default: 104)")
    parser.add_argument(
        "--error-rate",
        type=float,
        default=0.03,
        help="Fraction of rows (0-1) to deliberately corrupt, to exercise validate_data.py (default: 0.03)",
    )
    parser.add_argument("--seed", type=int, default=None, help="Random seed for reproducible output")
    parser.add_argument("--output", type=str, default=RAW_DATA_PATH, help="Path to write the raw CSV to")
    parser.add_argument(
        "--truth-output",
        type=str,
        default=TRUTH_DATA_PATH,
        help="Path to write each product's hidden true elasticity to",
    )
    args = parser.parse_args()

    rows, products = generate_dataset(args.products, args.weeks, args.error_rate, args.seed)
    write_csv(rows, CSV_FIELDS, args.output)
    write_csv(products, TRUTH_FIELDS, args.truth_output)
    print(
        f"Generated {len(rows)} weekly transaction rows for {len(products)} products "
        f"({round(args.error_rate * 100)}% target error rate) -> {args.output}"
    )
    print(f"Hidden true elasticities -> {args.truth_output}")


if __name__ == "__main__":
    main()

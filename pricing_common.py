"""Shared constants and data model for the SME price-setting tool.

Used by generate_data.py, validate_data.py, and analyze_pricing.py so all
three scripts agree on what a valid sales transaction looks like.
"""

from datetime import date

from pydantic import BaseModel, Field, model_validator

CATEGORIES = [
    "Coffee & Tea",
    "Bakery",
    "Snacks",
    "Household",
    "Personal Care",
    "Stationery",
    "Kitchenware",
    "Pet Supplies",
]

CHANNELS = ["In-store", "Online"]

RAW_DATA_PATH = "data/transactions_raw.csv"
VALIDATED_DATA_PATH = "data/transactions_validated.csv"
VALIDATION_REPORT_PATH = "data/validation_report.txt"
OUTPUT_WORKBOOK_PATH = "output/pricing_analysis.xlsx"
# Hidden per-product truth written by generate_data.py (optional input to analyze_pricing.py).
TRUTH_DATA_PATH = "data/products_truth.csv"

CSV_FIELDS = [
    "transaction_id",
    "week_start_date",
    "sku",
    "product_name",
    "category",
    "channel",
    "unit_cost",
    "unit_price",
    "units_sold",
    "revenue",
    "on_promo",
]

# Analysis thresholds -------------------------------------------------------

# Minimum number of distinct historical price points a product needs before
# we trust a regression on it at all.
MIN_PRICE_POINTS = 3
# Minimum number of weekly observations (with units_sold > 0) per product.
MIN_OBSERVATIONS = 12
# |t-statistic| above which the elasticity estimate is treated as statistically
# distinguishable from zero (roughly a 95% two-sided test for n > ~30).
MIN_ABS_T_STAT = 2.0
# Elasticity boundary between "inelastic" (safe to raise) and "elastic"
# (price-sensitive). Unit elasticity (-1) is where revenue is flat.
UNIT_ELASTICITY = -1.0

# Markdown depths (as fractions of current price) evaluated in the promo
# scenario table.
PROMO_DEPTHS = [0.05, 0.10, 0.15, 0.20, 0.25, 0.30]
# Price changes evaluated in the price-change simulation sheet.
PRICE_CHANGE_STEPS = [-0.20, -0.15, -0.10, -0.05, 0.05, 0.10, 0.15, 0.20]


class Transaction(BaseModel):
    """One product's aggregated weekly sales at a single price point."""

    transaction_id: str = Field(min_length=1)
    week_start_date: date
    sku: str = Field(min_length=1)
    product_name: str = Field(min_length=1)
    category: str = Field(min_length=1)
    channel: str = Field(min_length=1)
    unit_cost: float = Field(gt=0)
    unit_price: float = Field(gt=0)
    units_sold: int = Field(ge=0)
    revenue: float = Field(ge=0)
    on_promo: bool

    class Config:
        extra = "ignore"

    @model_validator(mode="after")
    def revenue_matches_price_times_units(self) -> "Transaction":
        expected = round(self.unit_price * self.units_sold, 2)
        if abs(self.revenue - expected) > 0.05:
            raise ValueError(
                f"revenue {self.revenue} does not equal unit_price x units_sold ({expected})"
            )
        return self

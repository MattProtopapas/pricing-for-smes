# Price Setting for SMEs

*How should a small company set its prices?*

Most SMEs price at cost-plus or copy a competitor, and plan holiday promotions with no
idea how their customers actually react to price. This is a small pipeline that generates
a fake sales history in which every product has been sold at several different prices,
validates it, estimates a **price elasticity per product** from that history, and turns
the result into three things a business owner can act on:

- a **"safe to raise"** list — inelastic products where a price increase adds contribution
- a **"price-sensitive"** list — elastic products where a well-judged discount can pay off
- a **promo scenario table** — for every product and markdown depth, the volume lift the
  discount *needs* to break even versus the lift the data says it will *actually* get

Everything is written into a single Excel workbook.

## The method

- **Deseasonalise first** — weekly units are divided by a calendar-month index (total units
  across all products), so the November/December peak isn't credited to the promotions that
  tend to run in the same weeks.
- **Log-log regression** — per product, `ln(units) = a + b · ln(price)` fitted by ordinary
  least squares on its weekly price/volume pairs. The slope `b` is the price elasticity:
  `b = -1.5` means a 1 % price cut lifts volume by about 1.5 %. Standard error, t-statistic
  and R² tell you how much to trust it.
- **Classification** — `-1 < b < 0` is inelastic → *Safe to raise*; `b < -1` is elastic →
  *Price-sensitive*. Products with too few price points / weeks are *Insufficient data*;
  those with `|t| < 2` are *Inconclusive*.
- **Contribution margin impact** — every scenario is judged on weekly contribution
  `units × (price − unit cost)`, never on revenue. A promo that grows revenue can still lose money.
- **Break-even lift for a discount** — the volume increase a markdown `d` needs to leave
  contribution unchanged: `(P − c) / (P·(1 − d) − c) − 1`. The expected lift from the
  elasticity is `(1 − d)^b − 1`. The promo pays only if expected lift exceeds break-even lift.
- **Markdown depth vs. lift** — the two lifts above are tabulated for 5 %–30 % markdowns per
  product, along with the deepest markdown that is still profitable.
- **Optimal margin (Lerner rule)** — for elastic products the contribution-maximising margin
  is `(P − c) / P = −1 / b`. Inelastic products have no interior optimum, so they are flagged
  for a bounded test increase instead.

## How it works

Three scripts, run in order, each consuming the previous one's output:

1. **`generate_data.py`** — uses [Faker](https://faker.readthedocs.io/) to create a set of
   products, each with a hidden *true* elasticity, then simulates ~2 years of weekly sales
   per product across 2–4 list-price changes plus random promotions (more frequent in
   Nov–Dec), with seasonality and noise. Writes `data/transactions_raw.csv`. A small
   percentage of rows are deliberately corrupted (negative prices, blank units, revenue that
   doesn't match price × units, an unparseable date, a duplicate transaction id) so the next
   step has real errors to catch. The hidden truth goes to `data/products_truth.csv` so you
   can check how well the analysis recovers it.
2. **`validate_data.py`** — checks every row in `data/transactions_raw.csv` against the schema
   in `pricing_common.py` (via [pydantic](https://docs.pydantic.dev/)), including a
   `revenue == unit_price × units_sold` consistency rule, drops invalid or duplicate rows,
   and writes:
   - `data/transactions_validated.csv` — the clean rows
   - `data/validation_report.txt` — what was rejected and why
3. **`analyze_pricing.py`** — reads `data/transactions_validated.csv`, fits the elasticity per
   product, and writes a new `output/pricing_analysis.xlsx` workbook with seven sheets:
   - **Input Data** — the validated transactions the analysis ran on
   - **Elasticity by Product** — slope, standard error, t-stat, R², price points seen, current
     price and margin, optimal margin, best price move within ±20 %, deepest profitable
     markdown, and a one-line recommendation. Rows are colour-coded by class. If
     `data/products_truth.csv` exists, a `true_elasticity` column is added for comparison.
   - **Safe to Raise** / **Price Sensitive** — the two action lists, filtered from the sheet above
   - **Promo Scenarios** — one row per product × markdown depth: promo price, break-even lift,
     expected lift, contribution before and after, and a *Profitable / Loses contribution /
     Below cost* verdict
   - **Price Change Simulation** — contribution impact of −20 % … +20 % price moves per product
   - **Method Notes** — plain-language explanation of every column, for the business reader

`pricing_common.py` holds the shared data model, file paths and analysis thresholds used by
all three scripts.

> **Example Results** : [sample_results.xlsx](./sample_results.xlsx)

## Setup

### 1. Get the code

```powershell
git clone <your-repo-url>
cd price-setting-for-smes
```

### 2. Create and activate a virtual environment

**Windows (PowerShell):**

```powershell
python -m venv .venv
.venv\Scripts\Activate.ps1
```

**macOS / Linux:**

```bash
python3 -m venv .venv
source .venv/bin/activate
```

### 3. Install dependencies

```powershell
pip install -r requirements.txt
```

Requires Python 3.10+. The regression is plain ordinary least squares written with the
standard library's `statistics` and `math` modules — no `numpy`, `scipy` or `pandas` needed.

## Usage

Run the three scripts in order from the project root, with the virtual environment
activated:

```powershell
python generate_data.py --products 40 --weeks 104 --seed 42
python validate_data.py
python analyze_pricing.py
```

Then open `output/pricing_analysis.xlsx`.

### Options

| Script | Flag | Default | Description |
| --- | --- | --- | --- |
| `generate_data.py` | `--products` | `40` | Number of products to generate |
| | `--weeks` | `104` | Weeks of sales history per product |
| | `--error-rate` | `0.03` | Fraction of rows to deliberately corrupt (`0` for a perfectly clean dataset) |
| | `--seed` | none | Random seed, for reproducible output |
| | `--output` | `data/transactions_raw.csv` | Where to write the raw data |
| | `--truth-output` | `data/products_truth.csv` | Where to write each product's hidden true elasticity |
| `validate_data.py` | `--input` | `data/transactions_raw.csv` | File to validate |
| | `--output` | `data/transactions_validated.csv` | Where to write the cleaned data |
| | `--report` | `data/validation_report.txt` | Where to write the rejection report |
| `analyze_pricing.py` | `--input` | `data/transactions_validated.csv` | File to analyze |
| | `--output` | `output/pricing_analysis.xlsx` | Where to write the results workbook |
| | `--truth` | `data/products_truth.csv` | Optional truth file; adds a comparison column if present |

Every script can also point at a custom input/output path, so you can drop in your own
real sales export (one row per product per week per channel, matching the columns in
`pricing_common.py`) and run it straight through `validate_data.py` and
`analyze_pricing.py`, skipping `generate_data.py`. The analysis thresholds (minimum price
points, minimum weeks, t-stat cut-off, markdown depths, simulated price steps) are
constants at the top of `pricing_common.py`.

## Reading the results

- A product is only worth acting on if it has **several distinct price points** and a
  **|t-stat| above 2**. Everything else is a candidate for a controlled price test, not a
  price change.
- On the **Safe to Raise** list, every simulated increase adds contribution. Start with a
  modest +5–10 % and re-measure rather than jumping to the +20 % edge of the simulation.
- On **Promo Scenarios**, look at `lift_gap_pct`: expected lift minus break-even lift. A
  positive gap is a discount that earns its keep; a negative gap is margin given away.
- Constant elasticity is an approximation. Competitor moves, stock-outs and display effects
  are ignored, and the model shouldn't be extrapolated beyond the `min_price`–`max_price`
  range actually observed.

## Project layout

```
price-setting-for-smes/
├── pricing_common.py       # shared schema, file paths, analysis thresholds
├── generate_data.py        # step 1: fake weekly sales with hidden true elasticities
├── validate_data.py        # step 2: schema validation / cleaning
├── analyze_pricing.py      # step 3: elasticity, margin impact, promo scenarios -> Excel
├── requirements.txt
├── sample_results.xlsx     # example output workbook
├── data/                   # generated CSVs + validation report (not committed)
└── output/                 # generated Excel workbook (not committed)
```

`data/` and `output/` are gitignored (apart from `.gitkeep` placeholders) since their
contents are reproducible by re-running the scripts — as is `.venv/`, so the virtual
environment itself is never pushed to the repo.


> **Disclaimer** — This software is provided for educational and demonstration purposes only. See the full [Disclaimer & EULA](./DISCLAIMER.md) for details.

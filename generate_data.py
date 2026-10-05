"""
generate_data.py
----------------
Creates a realistic *messy* retail-sales dataset (data/raw/sales_raw.csv).

Real-world data is never clean, so this script deliberately injects the
problems a data analyst meets every day:
  * missing values (customer, region, quantity, price, payment, date)
  * duplicate rows
  * inconsistent text (case, spaces, abbreviations, typos)
  * mixed date formats
  * numbers stored as text ("$1,200.50")
  * invalid values (negative quantity) and extreme outliers

Run:  python src/generate_data.py
"""
import random
from pathlib import Path

import numpy as np
import pandas as pd

SEED = 42
random.seed(SEED)
np.random.seed(SEED)

N_ROWS = 1500
OUT = Path(__file__).resolve().parents[1] / "data" / "raw" / "sales_raw.csv"

# product -> (category, base price)
PRODUCTS = {
    "Laptop": ("Electronics", 850.0),
    "Smartphone": ("Electronics", 600.0),
    "Headphones": ("Electronics", 120.0),
    "Office Chair": ("Furniture", 180.0),
    "Standing Desk": ("Furniture", 420.0),
    "Bookshelf": ("Furniture", 140.0),
    "Notebook Pack": ("Stationery", 12.0),
    "Pen Set": ("Stationery", 8.5),
    "Printer Paper": ("Stationery", 25.0),
    "Running Shoes": ("Apparel", 95.0),
    "Jacket": ("Apparel", 130.0),
    "T-Shirt": ("Apparel", 22.0),
}
REGIONS = ["North", "South", "East", "West"]
PAYMENTS = ["Credit Card", "Debit Card", "Cash", "PayPal", "Bank Transfer"]
FIRST = ["Aisha", "Liam", "Noah", "Emma", "Olivia", "Kasun", "Nimali", "Arjun",
         "Priya", "Sofia", "Lucas", "Mia", "Ethan", "Zara", "Ravi", "Chen"]
LAST = ["Perera", "Smith", "Fernando", "Khan", "Patel", "Garcia", "Silva",
        "Brown", "Wong", "Nguyen", "Johnson", "Ali", "Kumar", "Lopez"]

rows = []
dates = pd.date_range("2025-01-01", "2025-12-31")
for i in range(1, N_ROWS + 1):
    product = random.choice(list(PRODUCTS))
    category, base = PRODUCTS[product]
    qty = int(np.random.choice([1, 2, 3, 4, 5, 6, 8, 10], p=[.3, .22, .15, .1, .08, .06, .05, .04]))
    price = round(base * np.random.uniform(0.92, 1.08), 2)
    disc = int(np.random.choice([0, 5, 10, 15, 20], p=[.5, .2, .15, .1, .05]))
    total = round(qty * price * (1 - disc / 100), 2)
    rows.append({
        "order_id": f"ORD{1000 + i}",
        "order_date": random.choice(dates),
        "customer_name": f"{random.choice(FIRST)} {random.choice(LAST)}",
        "region": random.choice(REGIONS),
        "category": category,
        "product": product,
        "quantity": qty,
        "unit_price": price,
        "discount_pct": disc,
        "payment_method": random.choice(PAYMENTS),
        "total_amount": total,
    })

df = pd.DataFrame(rows)


def pick(n):
    return df.sample(n=n, random_state=random.randint(0, 10_000)).index


# ---- 1. inconsistent dates (stored as text in 4 different formats) --------
fmts = ["%Y-%m-%d", "%d/%m/%Y", "%b %d, %Y", "%d-%b-%Y"]
df["order_date"] = [d.strftime(random.choice(fmts)) for d in df["order_date"]]

# ---- 2. inconsistent text -------------------------------------------------
region_variants = {
    "North": ["north", "NORTH", " North ", "N.", "Nrth"],
    "South": ["south", "SOUTH", "South ", "S.", "South."],
    "East": ["east", "EAST", " East", "E.", "Est"],
    "West": ["west", "WEST", "West ", "W.", "Wst"],
}
for idx in pick(350):
    df.at[idx, "region"] = random.choice(region_variants[df.at[idx, "region"]])

cat_variants = {
    "Electronics": ["electronics", "ELECTRONICS", "Electroncs", "Electronic"],
    "Furniture": ["furniture", "FURNITURE", "Furnature", " Furniture "],
    "Stationery": ["stationery", "STATIONERY", "Stationary", "Stationery "],
    "Apparel": ["apparel", "APPAREL", "Aparel", "Clothing"],
}
for idx in pick(300):
    df.at[idx, "category"] = random.choice(cat_variants[df.at[idx, "category"]])

for idx in pick(200):
    df.at[idx, "customer_name"] = random.choice(
        [df.at[idx, "customer_name"].upper(), df.at[idx, "customer_name"].lower(),
         "  " + df.at[idx, "customer_name"] + "  "])

for idx in pick(150):
    df.at[idx, "payment_method"] = random.choice(
        ["credit card", "CREDIT CARD", "cc", "Debit", "paypal", "BANK TRANSFER", "cash "])

# ---- 3. numbers stored as text -------------------------------------------
df["unit_price"] = df["unit_price"].astype(object)
for idx in pick(120):
    df.at[idx, "unit_price"] = f"${df.at[idx, 'unit_price']:,.2f}"
df["total_amount"] = df["total_amount"].astype(object)
for idx in pick(100):
    df.at[idx, "total_amount"] = f"${df.at[idx, 'total_amount']:,.2f}"

# ---- 4. invalid values & outliers ----------------------------------------
for idx in pick(15):
    df.at[idx, "quantity"] = -abs(df.at[idx, "quantity"])       # negative qty
for idx in pick(10):
    df.at[idx, "quantity"] = random.choice([250, 500, 999])      # extreme outliers

# ---- 5. missing values ----------------------------------------------------
for col, n in [("customer_name", 90), ("region", 80), ("quantity", 70),
               ("unit_price", 60), ("payment_method", 75), ("category", 40),
               ("order_date", 12), ("discount_pct", 50), ("total_amount", 55)]:
    df.loc[pick(n), col] = np.nan

# ---- 6. duplicates (exact copies + repeated order ids) --------------------
exact_dupes = df.sample(n=60, random_state=1)
df = pd.concat([df, exact_dupes], ignore_index=True)
df = df.sample(frac=1, random_state=7).reset_index(drop=True)  # shuffle

OUT.parent.mkdir(parents=True, exist_ok=True)
df.to_csv(OUT, index=False)
print(f"Raw dataset written: {OUT}  ({len(df)} rows, {df.shape[1]} columns)")

"""
pipeline.py  --  Data Cleaning & Reporting Automation
=====================================================
One command cleans the raw sales data and regenerates every report:

    python src/pipeline.py                       # uses data/raw/sales_raw.csv
    python src/pipeline.py --input my_file.csv   # any file with the same columns

Outputs
  data/processed/sales_clean.csv        cleaned dataset
  reports/Sales_Report.xlsx             Excel report (formulas + native charts)
  reports/charts/*.png                  visual summaries
  reports/cleaning_log.csv / .txt       audit trail of every cleaning action

Stages: 1 Load & profile -> 2 Standardise text -> 3 Fix types & dates ->
        4 Remove duplicates -> 5 Handle invalid values & outliers ->
        6 Impute missing values -> 7 Validate -> 8 Report
"""
import argparse
import difflib
import logging
from datetime import datetime
from pathlib import Path

import matplotlib
matplotlib.use("Agg")
import matplotlib.pyplot as plt
import numpy as np
import pandas as pd
from openpyxl import Workbook
from openpyxl.chart import BarChart, LineChart, Reference
from openpyxl.styles import Alignment, Border, Font, PatternFill, Side
from openpyxl.utils import get_column_letter

ROOT = Path(__file__).resolve().parents[1]
RAW_DEFAULT = ROOT / "data" / "raw" / "sales_raw.csv"
CLEAN_OUT = ROOT / "data" / "processed" / "sales_clean.csv"
REPORT_DIR = ROOT / "reports"
CHART_DIR = REPORT_DIR / "charts"

VALID_REGIONS = ["North", "South", "East", "West"]
VALID_CATEGORIES = ["Electronics", "Furniture", "Stationery", "Apparel"]
VALID_PAYMENTS = ["Credit Card", "Debit Card", "Cash", "PayPal", "Bank Transfer"]
REGION_ALIASES = {"n": "North", "s": "South", "e": "East", "w": "West"}
CATEGORY_ALIASES = {"clothing": "Apparel", "electronic": "Electronics"}
PAYMENT_ALIASES = {"cc": "Credit Card", "debit": "Debit Card", "credit": "Credit Card"}
DATE_FORMATS = ["%Y-%m-%d", "%d/%m/%Y", "%b %d, %Y", "%d-%b-%Y"]
OUTLIER_IQR_MULT = 3.0   # only *extreme* outliers are capped

logging.basicConfig(level=logging.INFO, format="%(asctime)s | %(levelname)s | %(message)s",
                    datefmt="%H:%M:%S")
log = logging.getLogger("pipeline")
ACTIONS = []   # audit trail


def record(stage, issue, action, rows):
    ACTIONS.append({"Stage": stage, "Issue found": issue, "Action taken": action,
                    "Rows affected": int(rows)})
    log.info("%-22s %-40s -> %s (%d rows)", stage, issue, action, rows)


# ----------------------------------------------------------------------------
# helpers
# ----------------------------------------------------------------------------
def normalise_label(value, valid, aliases):
    """Map messy text to a valid label: exact -> alias -> fuzzy match."""
    if pd.isna(value):
        return np.nan
    key = str(value).strip().lower().rstrip(".")
    for v in valid:
        if key == v.lower():
            return v
    if key in aliases:
        return aliases[key]
    match = difflib.get_close_matches(key, [v.lower() for v in valid], n=1, cutoff=0.6)
    if match:
        return next(v for v in valid if v.lower() == match[0])
    return np.nan


def parse_date(value):
    if pd.isna(value):
        return pd.NaT
    for fmt in DATE_FORMATS:
        try:
            return datetime.strptime(str(value).strip(), fmt)
        except ValueError:
            continue
    return pd.NaT


def to_number(series):
    cleaned = series.astype(str).str.replace(r"[$,\s]", "", regex=True)
    return pd.to_numeric(cleaned.replace({"nan": np.nan, "": np.nan}), errors="coerce")


def profile(df):
    return {
        "rows": len(df),
        "missing_cells": int(df.isna().sum().sum()),
        "duplicate_rows": int(df.duplicated().sum()),
        "missing_by_col": df.isna().sum().to_dict(),
    }


# ----------------------------------------------------------------------------
# cleaning
# ----------------------------------------------------------------------------
def clean(raw):
    df = raw.copy()
    before = profile(raw)

    # --- 2. standardise text --------------------------------------------------
    for col in ["customer_name", "region", "category", "product", "payment_method", "order_id"]:
        df[col] = df[col].astype("string").str.strip().replace("", pd.NA)

    n = int((raw["customer_name"].dropna().astype(str) != raw["customer_name"].dropna().astype(str)
             .str.strip().str.title()).sum())
    df["customer_name"] = df["customer_name"].str.title()
    record("2 Standardise text", "Customer names: mixed case / extra spaces", "Trimmed + Title Case", n)

    for col, valid, aliases, label in [
        ("region", VALID_REGIONS, REGION_ALIASES, "Region"),
        ("category", VALID_CATEGORIES, CATEGORY_ALIASES, "Category"),
        ("payment_method", VALID_PAYMENTS, PAYMENT_ALIASES, "Payment method"),
    ]:
        original = df[col].copy()
        df[col] = original.map(lambda v: normalise_label(v, valid, aliases))
        changed = int(((original != df[col]) & original.notna() & df[col].notna()).sum())
        record("2 Standardise text", f"{label}: inconsistent spelling/case/abbreviations",
               "Mapped to standard labels (alias + fuzzy match)", changed)

    # --- 3. fix data types ----------------------------------------------------
    parsed = df["order_date"].map(parse_date)
    fmt_issue = int(raw["order_date"].dropna().astype(str).str.match(r"^\d{4}-\d{2}-\d{2}$").eq(False).sum())
    df["order_date"] = pd.to_datetime(parsed)
    record("3 Fix types", "Dates stored as text in 4 different formats", "Parsed to one datetime format", fmt_issue)

    for col in ["unit_price", "total_amount"]:
        txt = int(raw[col].astype(str).str.contains(r"[$,]").sum())
        df[col] = to_number(raw[col])
        record("3 Fix types", f"{col}: numbers stored as text with '$' / ','", "Converted to float", txt)
    for col in ["quantity", "discount_pct"]:
        df[col] = pd.to_numeric(df[col], errors="coerce")

    # --- 4. duplicates --------------------------------------------------------
    n = int(df.duplicated().sum())
    df = df.drop_duplicates()
    record("4 Duplicates", "Exact duplicate rows", "Dropped (kept first)", n)
    n = int(df["order_id"].duplicated().sum())
    df = df.drop_duplicates(subset="order_id", keep="first")
    record("4 Duplicates", "Repeated order_id after standardising", "Dropped (kept first)", n)

    # --- rows we cannot trust: no date -----------------------------------------
    n = int(df["order_date"].isna().sum())
    df = df.dropna(subset=["order_date"])
    record("4 Duplicates", "Missing order date (cannot be placed in time)", "Rows removed", n)

    # --- 5. invalid values & outliers ----------------------------------------
    neg = df["quantity"] < 0
    df.loc[neg, "quantity"] = df.loc[neg, "quantity"].abs()
    record("5 Invalid/outliers", "Negative quantity (sign entry error)", "Converted to absolute value", neg.sum())

    q1, q3 = df["quantity"].quantile([.25, .75])
    upper = q3 + OUTLIER_IQR_MULT * (q3 - q1)
    out = df["quantity"] > upper
    df.loc[out, "quantity"] = np.floor(upper)
    record("5 Invalid/outliers", f"Extreme quantity outliers (> {upper:.0f}, 3xIQR rule)",
           f"Capped at {int(np.floor(upper))}", out.sum())

    # --- 6. missing values ----------------------------------------------------
    prod_cat = df.dropna(subset=["category", "product"]).groupby("product")["category"] \
        .agg(lambda s: s.mode().iloc[0])
    m = df["category"].isna() & df["product"].notna()
    df.loc[m, "category"] = df.loc[m, "product"].map(prod_cat)
    record("6 Missing values", "Missing category", "Inferred from product", m.sum())

    for col, fill, why in [("customer_name", "Unknown", "label as Unknown"),
                           ("region", "Unknown", "label as Unknown"),
                           ("payment_method", "Unknown", "label as Unknown")]:
        n = int(df[col].isna().sum())
        df[col] = df[col].fillna(fill)
        record("6 Missing values", f"Missing {col}", f"Filled: {why}", n)

    for col in ["unit_price", "quantity"]:
        n = int(df[col].isna().sum())
        med = df.groupby("product")[col].transform("median")
        df[col] = df[col].fillna(med)
        if col == "quantity":
            df[col] = df[col].round().clip(lower=1)
        record("6 Missing values", f"Missing {col}", "Filled with median of same product", n)

    n = int(df["discount_pct"].isna().sum())
    df["discount_pct"] = df["discount_pct"].fillna(0)
    record("6 Missing values", "Missing discount_pct", "Assumed 0% (no discount)", n)

    stored = df["total_amount"].copy()
    df["total_amount"] = (df["quantity"] * df["unit_price"] * (1 - df["discount_pct"] / 100)).round(2)
    n_missing = int(stored.isna().sum())
    n_wrong = int(((stored - df["total_amount"]).abs() > 0.01).sum()) - n_missing
    record("6 Missing values", "Missing total_amount", "Recalculated = qty x price x (1-discount)", n_missing)
    record("6 Missing values", "Stored total disagrees after fixes (cap/abs/impute)",
           "Overwritten with recalculated value", n_wrong)

    # --- derived columns -------------------------------------------------------
    df["year_month"] = df["order_date"].dt.strftime("%Y-%m")
    df["order_date"] = df["order_date"].dt.date
    df["quantity"] = df["quantity"].astype(int)
    df["discount_pct"] = df["discount_pct"].astype(int)

    cols = ["order_id", "order_date", "year_month", "customer_name", "region", "category",
            "product", "quantity", "unit_price", "discount_pct", "payment_method", "total_amount"]
    df = df[cols].sort_values(["order_date", "order_id"]).reset_index(drop=True)

    # --- 7. validate ------------------------------------------------------------
    assert df.isna().sum().sum() == 0, "nulls remain"
    assert not df.duplicated().any() and df["order_id"].is_unique, "duplicates remain"
    assert (df["quantity"] > 0).all() and (df["unit_price"] > 0).all(), "invalid numbers remain"
    assert set(df["category"]) <= set(VALID_CATEGORIES), "bad categories"
    assert set(df["region"]) <= set(VALID_REGIONS + ["Unknown"]), "bad regions"
    log.info("Validation passed: dataset is clean (%d rows).", len(df))

    return df, before, profile(df)


# ----------------------------------------------------------------------------
# charts
# ----------------------------------------------------------------------------
PALETTE = ["#1F4E79", "#2E86C1", "#48C9B0", "#F5B041", "#E74C3C", "#8E44AD"]


def make_charts(df, before, after):
    CHART_DIR.mkdir(parents=True, exist_ok=True)
    plt.rcParams.update({"font.family": "DejaVu Sans", "axes.spines.top": False,
                         "axes.spines.right": False, "axes.grid": True, "grid.alpha": .25})
    kfmt = lambda x, _: f"${x/1000:,.0f}K"
    paths = {}

    def save(fig, name):
        p = CHART_DIR / f"{name}.png"
        fig.tight_layout()
        fig.savefig(p, dpi=150)
        plt.close(fig)
        paths[name] = p

    monthly = df.groupby("year_month")["total_amount"].sum()
    fig, ax = plt.subplots(figsize=(9, 4.2))
    ax.plot(monthly.index, monthly.values, marker="o", color=PALETTE[0], lw=2.5)
    ax.fill_between(monthly.index, monthly.values, alpha=.12, color=PALETTE[1])
    ax.set_title("Monthly Revenue Trend (2025)", fontweight="bold")
    ax.yaxis.set_major_formatter(kfmt)
    plt.setp(ax.get_xticklabels(), rotation=45)
    save(fig, "01_monthly_revenue")

    cat = df.groupby("category")["total_amount"].sum().sort_values()
    fig, ax = plt.subplots(figsize=(7, 4))
    bars = ax.barh(cat.index, cat.values, color=PALETTE[:len(cat)])
    ax.bar_label(bars, labels=[f"${v/1000:,.0f}K" for v in cat.values], padding=3)
    ax.set_title("Revenue by Category", fontweight="bold")
    ax.xaxis.set_major_formatter(kfmt)
    ax.grid(axis="y", visible=False)
    save(fig, "02_revenue_by_category")

    reg = df.groupby("region")["total_amount"].sum().sort_values(ascending=False)
    fig, ax = plt.subplots(figsize=(7, 4))
    bars = ax.bar(reg.index, reg.values, color=PALETTE[:len(reg)])
    ax.bar_label(bars, labels=[f"${v/1000:,.0f}K" for v in reg.values], padding=3)
    ax.set_title("Revenue by Region", fontweight="bold")
    ax.yaxis.set_major_formatter(kfmt)
    ax.grid(axis="x", visible=False)
    save(fig, "03_revenue_by_region")

    prod = df.groupby("product")["total_amount"].sum().sort_values().tail(8)
    fig, ax = plt.subplots(figsize=(7, 4.2))
    ax.barh(prod.index, prod.values, color=PALETTE[1])
    ax.set_title("Top 8 Products by Revenue", fontweight="bold")
    ax.xaxis.set_major_formatter(kfmt)
    ax.grid(axis="y", visible=False)
    save(fig, "04_top_products")

    pay = df["payment_method"].value_counts()
    fig, ax = plt.subplots(figsize=(6, 4.5))
    ax.pie(pay.values, labels=pay.index, autopct="%1.0f%%", startangle=90,
           colors=PALETTE + ["#95A5A6"], wedgeprops={"edgecolor": "white"})
    ax.set_title("Orders by Payment Method", fontweight="bold")
    save(fig, "05_payment_methods")

    heat = df.pivot_table(index="region", columns="category", values="total_amount", aggfunc="sum")
    fig, ax = plt.subplots(figsize=(8.5, 4.2))
    im = ax.imshow(heat.values, cmap="Blues", aspect="auto")
    ax.set_xticks(range(len(heat.columns)), heat.columns)
    ax.set_yticks(range(len(heat.index)), heat.index)
    for i in range(heat.shape[0]):
        for j in range(heat.shape[1]):
            ax.text(j, i, f"${heat.values[i, j]/1000:,.0f}K", ha="center", va="center",
                    color="white" if heat.values[i, j] > heat.values.mean() else "black", fontsize=9)
    ax.grid(False)
    ax.set_title("Revenue Heatmap: Region x Category", fontweight="bold")
    save(fig, "06_region_category_heatmap")

    cols = list(before["missing_by_col"])
    b = [before["missing_by_col"][c] for c in cols]
    a = [after["missing_by_col"][c] for c in cols]
    x = np.arange(len(cols))
    fig, ax = plt.subplots(figsize=(9, 4.2))
    ax.bar(x - .2, b, .4, label="Before cleaning", color=PALETTE[4])
    ax.bar(x + .2, a, .4, label="After cleaning", color=PALETTE[2])
    ax.set_xticks(x, cols, rotation=40, ha="right")
    ax.set_title("Missing Values per Column: Before vs After", fontweight="bold")
    ax.legend()
    save(fig, "07_missing_before_after")

    # one-page dashboard
    fig, axes = plt.subplots(2, 3, figsize=(16, 8.5))
    fig.suptitle("Sales Performance Dashboard (Cleaned Data)", fontsize=17, fontweight="bold")
    for ax, name in zip(axes.flat, ["01_monthly_revenue", "02_revenue_by_category",
                                    "03_revenue_by_region", "04_top_products",
                                    "05_payment_methods", "06_region_category_heatmap"]):
        ax.imshow(plt.imread(paths[name]))
        ax.axis("off")
    fig.tight_layout()
    fig.savefig(CHART_DIR / "00_dashboard.png", dpi=130)
    plt.close(fig)
    log.info("Charts saved to %s", CHART_DIR)


# ----------------------------------------------------------------------------
# Excel report
# ----------------------------------------------------------------------------
FONT = "Arial"
HDR_FILL = PatternFill("solid", fgColor="1F4E79")
ALT_FILL = PatternFill("solid", fgColor="EAF2F8")
THIN = Side(style="thin", color="BFC9D1")
BORDER = Border(top=THIN, bottom=THIN, left=THIN, right=THIN)


def style_header(ws, row, ncols, start_col=1):
    for c in range(start_col, start_col + ncols):
        cell = ws.cell(row=row, column=c)
        cell.font = Font(name=FONT, bold=True, color="FFFFFF")
        cell.fill = HDR_FILL
        cell.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)
        cell.border = BORDER


def style_body(ws, r1, r2, ncols, start_col=1):
    for r in range(r1, r2 + 1):
        for c in range(start_col, start_col + ncols):
            cell = ws.cell(row=r, column=c)
            cell.font = Font(name=FONT, size=10)
            cell.border = BORDER
            if (r - r1) % 2 == 1:
                cell.fill = ALT_FILL


def title(ws, text, sub=None):
    ws["A1"] = text
    ws["A1"].font = Font(name=FONT, size=16, bold=True, color="1F4E79")
    if sub:
        ws["A2"] = sub
        ws["A2"].font = Font(name=FONT, size=10, italic=True, color="666666")
    ws.sheet_view.showGridLines = False


def widths(ws, mapping):
    for col, w in mapping.items():
        ws.column_dimensions[col].width = w


def build_excel(df, before, after, path):
    wb = Workbook()
    n = len(df)
    last = n + 1
    # column letters in the Clean Data sheet
    cols = list(df.columns)
    L = {c: get_column_letter(i + 1) for i, c in enumerate(cols)}
    rng = lambda c: f"'Clean Data'!${L[c]}$2:${L[c]}${last}"

    # ---------------- Clean Data ----------------
    wsd = wb.active
    wsd.title = "Clean Data"
    wsd.append(cols)
    for row in df.itertuples(index=False):
        wsd.append(list(row))
    style_header(wsd, 1, len(cols))
    for r in range(2, last + 1):
        for c in range(1, len(cols) + 1):
            wsd.cell(row=r, column=c).font = Font(name=FONT, size=10)
        wsd.cell(row=r, column=2).number_format = "yyyy-mm-dd"
        wsd.cell(row=r, column=9).number_format = "$#,##0.00"
        wsd.cell(row=r, column=12).number_format = "$#,##0.00"
    wsd.freeze_panes = "A2"
    wsd.auto_filter.ref = f"A1:{get_column_letter(len(cols))}{last}"
    widths(wsd, {"A": 11, "B": 12, "C": 11, "D": 20, "E": 9, "F": 13, "G": 15,
                 "H": 9, "I": 11, "J": 12, "K": 15, "L": 13})

    # ---------------- By Region / By Category (same layout) ----------------
    def group_sheet(name, key, label):
        ws = wb.create_sheet(name)
        title(ws, f"Sales by {label}", "All figures are live formulas over the 'Clean Data' sheet.")
        ws.append([])
        hdr = [label, "Orders", "Units Sold", "Revenue ($)", "Share of Revenue", "Avg Order Value ($)"]
        ws.append(hdr)
        style_header(ws, 4, len(hdr))
        order = df.groupby(key)["total_amount"].sum().sort_values(ascending=False).index
        r0 = 5
        for i, g in enumerate(order):
            r = r0 + i
            ws.cell(row=r, column=1, value=g)
            ws.cell(row=r, column=2, value=f"=COUNTIF({rng(key)},A{r})")
            ws.cell(row=r, column=3, value=f"=SUMIFS({rng('quantity')},{rng(key)},A{r})")
            ws.cell(row=r, column=4, value=f"=SUMIFS({rng('total_amount')},{rng(key)},A{r})")
            ws.cell(row=r, column=5, value=f"=IFERROR(D{r}/$D${r0 + len(order)},0)")
            ws.cell(row=r, column=6, value=f"=IFERROR(D{r}/B{r},0)")
        rt = r0 + len(order)
        ws.cell(row=rt, column=1, value="Total")
        for c, letter in [(2, "B"), (3, "C"), (4, "D")]:
            ws.cell(row=rt, column=c, value=f"=SUM({letter}{r0}:{letter}{rt - 1})")
        ws.cell(row=rt, column=5, value=f"=SUM(E{r0}:E{rt - 1})")
        ws.cell(row=rt, column=6, value=f"=IFERROR(D{rt}/B{rt},0)")
        style_body(ws, r0, rt, len(hdr))
        for c in range(1, len(hdr) + 1):
            ws.cell(row=rt, column=c).font = Font(name=FONT, size=10, bold=True)
        for r in range(r0, rt + 1):
            ws.cell(row=r, column=2).number_format = "#,##0"
            ws.cell(row=r, column=3).number_format = "#,##0"
            ws.cell(row=r, column=4).number_format = "$#,##0"
            ws.cell(row=r, column=5).number_format = "0.0%"
            ws.cell(row=r, column=6).number_format = "$#,##0.00"
        widths(ws, {"A": 16, "B": 10, "C": 12, "D": 14, "E": 16, "F": 20})
        ch = BarChart()
        ch.type = "col"
        ch.title = f"Revenue by {label}"
        ch.y_axis.title = "Revenue ($)"
        ch.add_data(Reference(ws, min_col=4, min_row=4, max_row=rt - 1), titles_from_data=True)
        ch.set_categories(Reference(ws, min_col=1, min_row=r0, max_row=rt - 1))
        ch.legend = None
        ch.height, ch.width = 8, 15
        ws.add_chart(ch, f"H4")
        return ws, r0, rt

    ws_reg, reg_r0, reg_rt = group_sheet("By Region", "region", "Region")
    ws_cat, cat_r0, cat_rt = group_sheet("By Category", "category", "Category")

    # ---------------- Monthly Trend ----------------
    wsm = wb.create_sheet("Monthly Trend")
    title(wsm, "Monthly Revenue Trend", "Live formulas over the 'Clean Data' sheet.")
    wsm.append([])
    hdr = ["Month", "Orders", "Revenue ($)", "MoM Growth"]
    wsm.append(hdr)
    style_header(wsm, 4, len(hdr))
    months = sorted(df["year_month"].unique())
    for i, m in enumerate(months):
        r = 5 + i
        wsm.cell(row=r, column=1, value=m)
        wsm.cell(row=r, column=2, value=f"=COUNTIF({rng('year_month')},A{r})")
        wsm.cell(row=r, column=3, value=f"=SUMIFS({rng('total_amount')},{rng('year_month')},A{r})")
        wsm.cell(row=r, column=4, value=f'=IFERROR(C{r}/C{r - 1}-1,"")' if i else "")
    mend = 4 + len(months)
    style_body(wsm, 5, mend, len(hdr))
    for r in range(5, mend + 1):
        wsm.cell(row=r, column=2).number_format = "#,##0"
        wsm.cell(row=r, column=3).number_format = "$#,##0"
        wsm.cell(row=r, column=4).number_format = "0.0%"
    widths(wsm, {"A": 12, "B": 10, "C": 14, "D": 14})
    lc = LineChart()
    lc.title = "Monthly Revenue"
    lc.y_axis.title = "Revenue ($)"
    lc.add_data(Reference(wsm, min_col=3, min_row=4, max_row=mend), titles_from_data=True)
    lc.set_categories(Reference(wsm, min_col=1, min_row=5, max_row=mend))
    lc.legend = None
    lc.height, lc.width = 8, 16
    wsm.add_chart(lc, "F4")

    # ---------------- Top Products ----------------
    wsp = wb.create_sheet("Top Products")
    title(wsp, "Product Performance", "Sorted by revenue; live formulas over 'Clean Data'.")
    wsp.append([])
    hdr = ["Product", "Category", "Units Sold", "Revenue ($)"]
    wsp.append(hdr)
    style_header(wsp, 4, len(hdr))
    prod = df.groupby(["product", "category"])["total_amount"].sum().sort_values(ascending=False).reset_index()
    for i, row in prod.iterrows():
        r = 5 + i
        wsp.cell(row=r, column=1, value=row["product"])
        wsp.cell(row=r, column=2, value=row["category"])
        wsp.cell(row=r, column=3, value=f"=SUMIFS({rng('quantity')},{rng('product')},A{r})")
        wsp.cell(row=r, column=4, value=f"=SUMIFS({rng('total_amount')},{rng('product')},A{r})")
    pend = 4 + len(prod)
    style_body(wsp, 5, pend, len(hdr))
    for r in range(5, pend + 1):
        wsp.cell(row=r, column=3).number_format = "#,##0"
        wsp.cell(row=r, column=4).number_format = "$#,##0"
    widths(wsp, {"A": 18, "B": 14, "C": 12, "D": 14})
    pc = BarChart()
    pc.type = "bar"
    pc.title = "Revenue by Product"
    pc.add_data(Reference(wsp, min_col=4, min_row=4, max_row=pend), titles_from_data=True)
    pc.set_categories(Reference(wsp, min_col=1, min_row=5, max_row=pend))
    pc.legend = None
    pc.height, pc.width = 9, 15
    wsp.add_chart(pc, "F4")

    # ---------------- Data Quality ----------------
    wsq = wb.create_sheet("Data Quality")
    title(wsq, "Data Quality: Before vs After Cleaning",
          "Static snapshot produced by the pipeline at run time.")
    wsq.append([])
    hdr = ["Metric", "Before", "After"]
    wsq.append(hdr)
    style_header(wsq, 4, 3)
    metrics = [("Total rows", before["rows"], after["rows"]),
               ("Missing cells", before["missing_cells"], after["missing_cells"]),
               ("Duplicate rows", before["duplicate_rows"], after["duplicate_rows"])]
    for i, (m, b, a) in enumerate(metrics):
        wsq.cell(row=5 + i, column=1, value=m)
        wsq.cell(row=5 + i, column=2, value=b)
        wsq.cell(row=5 + i, column=3, value=a)
    style_body(wsq, 5, 7, 3)
    wsq["A9"] = "Missing values by column"
    wsq["A9"].font = Font(name=FONT, bold=True, size=11, color="1F4E79")
    wsq.append([])
    for c, h in enumerate(["Column", "Before", "After"], 1):
        wsq.cell(row=10, column=c, value=h)
    style_header(wsq, 10, 3)
    for i, col in enumerate(before["missing_by_col"]):
        wsq.cell(row=11 + i, column=1, value=col)
        wsq.cell(row=11 + i, column=2, value=before["missing_by_col"][col])
        wsq.cell(row=11 + i, column=3, value=after["missing_by_col"][col])
    qend = 10 + len(before["missing_by_col"])
    style_body(wsq, 11, qend, 3)
    widths(wsq, {"A": 22, "B": 12, "C": 12})

    # ---------------- Cleaning Log ----------------
    wsl = wb.create_sheet("Cleaning Log")
    title(wsl, "Cleaning Log (audit trail)", "Every transformation applied to the raw data.")
    wsl.append([])
    hdr = ["Stage", "Issue found", "Action taken", "Rows affected"]
    wsl.append(hdr)
    style_header(wsl, 4, 4)
    for i, a in enumerate(ACTIONS):
        for c, k in enumerate(hdr, 1):
            wsl.cell(row=5 + i, column=c, value=a[k])
    lend = 4 + len(ACTIONS)
    style_body(wsl, 5, lend, 4)
    for r in range(5, lend + 1):
        wsl.cell(row=r, column=4).alignment = Alignment(horizontal="center")
    widths(wsl, {"A": 22, "B": 58, "C": 50, "D": 14})

    # ---------------- Summary (first sheet) ----------------
    wss = wb.create_sheet("Summary", 0)
    title(wss, "Sales Report: Executive Summary",
          f"Generated automatically on {datetime.now():%d %b %Y %H:%M} by src/pipeline.py")
    wss.append([])
    kpis = [
        ("Total Revenue ($)", f"=SUM({rng('total_amount')})", "$#,##0"),
        ("Total Orders", f"=COUNTA({rng('order_id')})", "#,##0"),
        ("Units Sold", f"=SUM({rng('quantity')})", "#,##0"),
        ("Average Order Value ($)", "=IFERROR(B5/B6,0)", "$#,##0.00"),
        ("Average Discount (%)", f"=AVERAGE({rng('discount_pct')})/100", "0.0%"),
        ("Top Region (by revenue)",
         f"=INDEX('By Region'!$A${reg_r0}:$A${reg_rt - 1},MATCH(MAX('By Region'!$D${reg_r0}:$D${reg_rt - 1}),'By Region'!$D${reg_r0}:$D${reg_rt - 1},0))", "@"),
        ("Top Category (by revenue)",
         f"=INDEX('By Category'!$A${cat_r0}:$A${cat_rt - 1},MATCH(MAX('By Category'!$D${cat_r0}:$D${cat_rt - 1}),'By Category'!$D${cat_r0}:$D${cat_rt - 1},0))", "@"),
        ("Best Month (by revenue)",
         f"=INDEX('Monthly Trend'!$A$5:$A${mend},MATCH(MAX('Monthly Trend'!$C$5:$C${mend}),'Monthly Trend'!$C$5:$C${mend},0))", "@"),
    ]
    wss.append(["Key Performance Indicator", "Value"])
    style_header(wss, 4, 2)
    for i, (k, f, fmt) in enumerate(kpis):
        r = 5 + i
        wss.cell(row=r, column=1, value=k)
        c = wss.cell(row=r, column=2, value=f)
        c.number_format = fmt
        c.alignment = Alignment(horizontal="right")
    style_body(wss, 5, 4 + len(kpis), 2)
    for r in range(5, 5 + len(kpis)):
        wss.cell(row=r, column=1).font = Font(name=FONT, size=10, bold=True)
        wss.cell(row=r, column=2).font = Font(name=FONT, size=11, bold=True, color="1F4E79")

    r = 6 + len(kpis)
    wss.cell(row=r, column=1, value="Data cleaning snapshot").font = Font(name=FONT, bold=True, size=12, color="1F4E79")
    snap = [("Raw rows received", before["rows"]),
            ("Clean rows delivered", after["rows"]),
            ("Missing cells fixed", before["missing_cells"]),
            ("Duplicate rows removed", before["duplicate_rows"] + sum(
                a["Rows affected"] for a in ACTIONS if a["Issue found"].startswith("Repeated order_id")))]
    for i, (k, v) in enumerate(snap):
        wss.cell(row=r + 1 + i, column=1, value=k)
        wss.cell(row=r + 1 + i, column=2, value=v).number_format = "#,##0"
    style_body(wss, r + 1, r + len(snap), 2)

    r2 = r + len(snap) + 2
    wss.cell(row=r2, column=1, value="Sheets in this workbook").font = Font(name=FONT, bold=True, size=12, color="1F4E79")
    guide = [("By Region / By Category", "Revenue, orders, share and AOV with charts"),
             ("Monthly Trend", "Revenue by month with growth %"),
             ("Top Products", "Product ranking"),
             ("Data Quality", "Before vs after cleaning"),
             ("Cleaning Log", "Audit trail of every fix"),
             ("Clean Data", "The cleaned dataset (source of all formulas)")]
    for i, (k, v) in enumerate(guide):
        wss.cell(row=r2 + 1 + i, column=1, value=k).font = Font(name=FONT, size=10, bold=True)
        wss.cell(row=r2 + 1 + i, column=2, value=v).font = Font(name=FONT, size=10)
    widths(wss, {"A": 30, "B": 48})

    wb.move_sheet("Clean Data", offset=len(wb.sheetnames) - 1 - wb.sheetnames.index("Clean Data"))
    wb.save(path)
    log.info("Excel report saved: %s", path)


# ----------------------------------------------------------------------------
# main
# ----------------------------------------------------------------------------
def main():
    ap = argparse.ArgumentParser(description="Clean raw sales data and regenerate reports.")
    ap.add_argument("--input", default=str(RAW_DEFAULT), help="path to raw CSV")
    args = ap.parse_args()

    log.info("=== Data Cleaning & Reporting Automation ===")
    raw = pd.read_csv(args.input, dtype=object)
    for col in ["quantity", "discount_pct"]:
        raw[col] = pd.to_numeric(raw[col], errors="coerce")
    log.info("Loaded %d rows x %d columns from %s", *raw.shape, args.input)

    df, before, after = clean(raw)

    CLEAN_OUT.parent.mkdir(parents=True, exist_ok=True)
    REPORT_DIR.mkdir(parents=True, exist_ok=True)
    df.to_csv(CLEAN_OUT, index=False)
    log.info("Clean dataset saved: %s", CLEAN_OUT)

    log_df = pd.DataFrame(ACTIONS)
    log_df.to_csv(REPORT_DIR / "cleaning_log.csv", index=False)
    with open(REPORT_DIR / "cleaning_log.txt", "w") as f:
        f.write(f"Cleaning run: {datetime.now():%Y-%m-%d %H:%M:%S}\n")
        f.write(f"Rows in: {before['rows']}   Rows out: {after['rows']}\n")
        f.write(f"Missing cells before: {before['missing_cells']}   after: {after['missing_cells']}\n\n")
        f.write(log_df.to_string(index=False))

    make_charts(df, before, after)
    build_excel(df, before, after, REPORT_DIR / "Sales_Report.xlsx")
    log.info("Done. Rows: %d -> %d | Missing cells: %d -> %d",
             before["rows"], after["rows"], before["missing_cells"], after["missing_cells"])


if __name__ == "__main__":
    main()

# Data Cleaning & Reporting Automation

**Goal:** automate the cleaning of a messy sales dataset and generate reports and visual summaries
with a single command, using Python and Excel.

| Requirement | How this project delivers it |
|---|---|
| Use Python, Excel, or Power BI | Python (pandas, matplotlib) for cleaning/charts, Excel (openpyxl) for the report |
| Handle missing values | 9 columns with gaps: imputed by product median, inferred from product, or labelled `Unknown` |
| Handle duplicates | 60 exact duplicates removed, plus an `order_id` uniqueness check |
| Handle inconsistent data | Region, category, payment, names, dates, and prices normalised to one standard |
| Generate automated reports | `Sales_Report.xlsx`: 8 sheets, 122 live formulas, native charts |
| Generate visual summaries | 8 PNG charts plus a one-page dashboard |

## Dataset
`data/raw/sales_raw.csv`: 1,560 retail orders for 2025 (11 columns): order, date, customer, region,
category, product, quantity, unit price, discount, payment method, total.
The data is synthetic (built by `src/generate_data.py`) with problems injected on purpose.
To use real data, drop in a CSV with the same columns and run `python src/pipeline.py --input yourfile.csv`.

## Pipeline (`src/pipeline.py`)
1. **Load and profile**: count rows, missing cells, duplicates (the "before" snapshot).
2. **Standardise text**: trim spaces, Title Case names, map region/category/payment to valid labels
   (exact match, then alias table, then fuzzy match with `difflib`), e.g. `Nrth`, `N.`, `NORTH` become `North`.
3. **Fix types**: four date formats parsed to one; `"$1,200.50"` converted to a number.
4. **Remove duplicates**: exact duplicates, then repeated `order_id`; rows with no date are dropped.
5. **Invalid values and outliers**: negative quantities treated as sign errors (absolute value);
   extreme quantities (250, 500, 999) capped using the 3xIQR rule.
6. **Missing values**: category inferred from product; price and quantity filled with the product median;
   customer, region, and payment labelled `Unknown`; discount assumed 0%; `total_amount` recalculated as
   `quantity x unit_price x (1 - discount)`.
7. **Validate**: assertions confirm no nulls, no duplicates, no invalid numbers, and only valid labels.
8. **Report**: save clean CSV, cleaning log, charts, and the Excel workbook.

## Results
| Metric | Before | After |
|---|---|---|
| Rows | 1,560 | 1,488 |
| Missing cells | 554 | 0 |
| Duplicate rows | 60 | 0 |

Every action and the number of rows it affected is recorded in `reports/cleaning_log.csv` and in the
**Cleaning Log** sheet of the workbook.

## Key insights
* Total revenue is about **$943K** from 1,488 orders (average order value about $634).
* **Electronics** earns about 59% of revenue; **Laptop** and **Smartphone** are the top two products.
* **West** is the strongest region; **December** is the best month and **July** shows a dip.
* About $67K of revenue has an `Unknown` region, so capturing region at order entry would improve analysis.

## Assumptions and limitations
* Imputing with medians keeps totals realistic but is an estimate, not the true original value.
* Negative quantities are treated as typos; in a real business they might be returns and need a separate flag.
* Missing region/customer/payment are labelled `Unknown` rather than guessed.

## Automation
Run `./run_all.sh` (or `python src/pipeline.py`) any time new raw data arrives. To run it on a schedule:
* Windows: Task Scheduler running `python src\pipeline.py`
* Linux/macOS: cron entry such as `0 8 * * 1 cd /path/to/project && python src/pipeline.py`

The generated `reports/Sales_Report.xlsx` can also be loaded into **Power BI** (Get Data > Excel, use the `Clean Data` sheet).

## Project structure
```
data/raw/sales_raw.csv          messy input
data/processed/sales_clean.csv  cleaned output
src/generate_data.py            builds the messy dataset
src/pipeline.py                 cleaning + reporting pipeline
reports/Sales_Report.xlsx       Excel report
reports/charts/                 PNG visual summaries
reports/cleaning_log.csv|txt    audit trail
docs/Project_Report.docx        written project report (objective, method, results, insights)
```

## How to run
```
pip install -r requirements.txt
python src/generate_data.py
python src/pipeline.py
```

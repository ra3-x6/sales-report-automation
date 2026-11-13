# Sales Report Automation

[![Python checks](https://github.com/ra3-x6/sales-report-automation/actions/workflows/ci.yml/badge.svg)](https://github.com/ra3-x6/sales-report-automation/actions/workflows/ci.yml)

Turn one or more order CSV exports into an Excel report with a daily chart, product totals, source order lines, and a sheet explaining rejected rows. Run it on the included synthetic data or a CSV exported from the companion [Orders and Inventory API](https://github.com/ra3-x6/orders-inventory-api).

## Run the example

From this directory in PowerShell, use Python 3.11 or later.

```powershell
python -m venv .venv
.\.venv\Scripts\python.exe -m pip install -e ".[dev]"
.\.venv\Scripts\sales-report.exe examples/orders-a.csv examples/orders-b.csv -o output/sales.xlsx
```

On macOS or Linux, use `.venv/bin/python` and `.venv/bin/sales-report`. The command can also be run with `python -m sales_report.cli` inside the project environment.

Expected results

| Metric | Value |
| --- | ---: |
| Completed orders | 2 |
| Cancelled orders | 1 |
| Pending orders | 1 |
| Units sold | 5 |
| Completed revenue | EGP 265.50 |
| Accepted order lines | 5 |
| Duplicate rows skipped | 1 |
| Rejected rows | 2 |

## Use an API export

Keep the two project directories next to each other. Run the Orders and Inventory API demo against a fresh database, then run the reporting CLI from this directory.

```powershell
.\.venv\Scripts\sales-report.exe ../orders-inventory-api/output/orders.csv -o output/api-sales.xlsx
```

On macOS or Linux, use `.venv/bin/sales-report` as the executable. The demo report shows EGP 400.00 in revenue from two completed orders and lists the cancelled order in **Order Lines**.

## Input contract

Use UTF-8 CSV with exactly these named columns. Header order can vary.

```csv
order_id,line_id,ordered_at,status,sku,product_name,quantity,unit_price_minor,currency
demo-101,1,2026-10-02T10:00:00Z,completed,BURGER,House Burger,2,9500,EGP
```

- One row represents one order line. The pair `order_id` and `line_id` uniquely identifies a line. Identical rows count once. Conflicts within an order or a failed row validation exclude the entire known order and record the reason.
- Use `pending`, `completed`, or `cancelled`. Revenue includes completed orders. Daily totals use the selected reporting timezone and the timestamp at which each order was placed.
- Quantities and prices are integers. Money stays in minor units until it is written into Excel. The example defaults to EGP and Africa/Cairo. Use `--currency USD` or `--currency EUR` for a single-currency export and `--timezone UTC` for UTC day boundaries.
- Each run writes a complete report. Unreadable CSV files or an invalid header exit with an error and preserve an existing output. Rows with invalid values are explained in the workbook.

## Workbook output

**Summary** contains the main counts and sales chart. **Daily Sales** and **Product Sales** break down revenue. **Order Lines** preserves accepted rows and **Rejected Rows** records the file, CSV row number, and reason.

## Run the checks

```powershell
.\.venv\Scripts\python.exe -m pytest -q
.\.venv\Scripts\python.exe -m ruff check .
```

The tests check the workbook totals against the sample exports, conflicting duplicate lines, timezone boundaries, empty exports, and handling imported text as text in Excel.

[Open the sample workbook](examples/sales-preview.xlsx) to see the report generated from both example exports.

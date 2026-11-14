# Amazon.in Order + Payment Reconciliation Automation Tool

[![Python 3.11](https://img.shields.io/badge/Python-3.11-blue.svg)](https://www.python.org/)
[![Tests](https://img.shields.io/badge/pytest-32%20passed-brightgreen.svg)]()
[![License](https://img.shields.io/badge/License-Proprietary-darkred.svg)]()

A high-precision financial reconciliation automation tool tailored for Amazon India (`Amazon.in`) apparel sellers operating merchant-fulfilled (FBM via Easy Ship) storefronts with size-based SKUs (e.g. `AA(ARN)BLACK_173-M`).

---

## 1. What the Tool Does

Each time you run the tool, you provide two files (or simply drop them into `/data/inbox`):
1. **Order Report** (`.txt` / `.tsv` tab-separated file from Seller Central)
2. **Payment Report** (`CustomUnifiedTransaction.csv` from Amazon Payments)

The tool joins them on Amazon Order ID and generates **ONE** comprehensive, audit-ready multi-tab Excel workbook:
- **`Order Pivot`**: Main reconciliation table with delivery status, autofilters, and formula Grand Totals.
- **`Status Summary`**: Executive dashboard with KPI cards and delivery status return rate breakdown.
- **`SKU Performance`**: Granular matrix decomposing apparel SKUs by Design Base and Size with return rates.
- **`Non-Order Rows`**: Transfers, Debt adjustments, Service Fees, and Closing Amazon Balance Roll-Forward.
- **`Exceptions`**: Audit trail of every discrepancy with reason codes and severity levels.

---

## 2. Delivery Status Determination Rules

Delivery status is determined using exact `Decimal` precision based on the configured basis (default: `net_ex_shipping`):

| Condition | `net_ex_shipping` Basis (Default) | `total_net` Basis |
| :--- | :---: | :---: |
| Value > 0 | **`Delivered`** | **`Delivered`** |
| Value == 0 | **`RTO`** (Return to Origin) | **`RTO`** |
| Value < 0 | **`RVP`** (Customer Return) | **`RVP`** |
| In Order report only, status `Cancelled` | **`Cancelled`** | **`Cancelled`** |
| In Order report only, shipped, not yet in payments | **`Not settled`** (`SHIPPED_NOT_SETTLED`) | **`Not settled`** |
| In Payment only (not yet in order report) | **`Delivered` / `RVP`** (`SETTLED_NOT_IN_ORDERS`) | **`Delivered` / `RVP`** |

*Note: For replacement/warranty orders with `Order = 0.00` and `Shipping = -64.90`, `net_ex_shipping` resolves to `0.00` (`RTO`), while `total_net` resolves to `-64.90` (`RVP`).*

---

## 3. Directory Layout

```text
amazon_recon/
├── config/
│   └── config.yaml             <- Application configuration and tolerances
├── data/
│   ├── inbox/                  <- Drop your order (.txt) and payment (.csv) files here
│   ├── raw/
│   │   ├── orders/             <- Untouched archived copies of raw order reports
│   │   └── payments/           <- Untouched archived copies of raw payment reports
│   └── processed/              <- Archived files moved from inbox after processing
├── reports/                    <- Generated Excel reconciliation reports
├── logs/                       <- Rotating execution logs (logs/reconciliation.log)
├── src/
│   ├── config.py               <- AppConfig loader & validation
│   ├── logger.py               <- Dual stream & rotating file logger
│   ├── orders.py               <- Order report ingestion, deduplication & summary
│   ├── payments.py             <- Payment report ingestion, header finder & pivot
│   ├── reconcile.py            <- Full outer join & delivery status engine
│   ├── pipeline.py             <- Multi-file orchestration & console dashboard
│   └── reporting/
│       └── excel_export.py     <- 5-Sheet Excel workbook generator
├── tests/                      <- Complete pytest suite (32 tests passing)
├── main.py                     <- CLI entrypoint
└── run_reconciliation.bat      <- 1-Click Windows execution script
```

---

## 4. How to Run

### Method A: 1-Click Desktop Runner
Double-click `run_reconciliation.bat` in the project root. It will automatically:
1. Scan `/data/inbox` for any dropped order `.txt` and payment `.csv` files.
2. Auto-detect file types based on headers.
3. Archive raw files untouched to `/data/raw/`.
4. Reconcile orders and payments.
5. Save the Excel report to `/reports/order_payment_recon_<date>.xlsx`.
6. Display a summary dashboard on screen.

### Method B: Python CLI Commands

```powershell
# 1. Automatic mode (scans /data/inbox)
python main.py --auto

# 2. Explicit file paths
python main.py --orders "path/to/orders.txt" --payments "path/to/payments.csv"

# 3. Date range filter (IST dates YYYY-MM-DD)
python main.py --auto --from 2026-08-01 --to 2026-10-31

# 4. Evaluation basis override (net_ex_shipping vs total_net)
python main.py --auto --basis total_net

# 5. Automatically open the Excel report on completion
python main.py --auto --open
```

---

## 5. Running Tests

```powershell
Set-Location "C:\Users\knitk\.gemini\antigravity\scratch\amazon_recon"
.\.venv\Scripts\pytest.exe -v tests/test_setup.py tests/test_orders.py tests/test_payments.py tests/test_reconcile.py tests/test_excel_export.py tests/test_cli_pipeline.py
```
*Current test status: **32 passed in ~2.6s**.*

---

## 6. Engineering Standards Enforced
- **Zero Raw File Mutation**: Files are copied untouched to `/data/raw/` before any parsing.
- **Strict Decimal Arithmetic**: All monetary calculations use `Decimal` with 2-decimal rounding; floating-point money is prohibited.
- **Idempotency**: Hash checks on file level (`MD5`) and payment row level (`row_hash`) prevent double-counting across overlapping runs.
- **Zero Dropped Rows**: Every unexpected or non-matching row is retained and logged as an audit exception with severity and reason codes.
- **Buyer Address Privacy**: Only `ship-state` and `ship-postal-code` are retained; personal buyer information is discarded.

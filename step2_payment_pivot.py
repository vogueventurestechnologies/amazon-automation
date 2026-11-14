"""Step 2: Amazon Payment Report Pivot for Order ID and Total.

Workflow:
1. Reads Amazon Payment Report (CustomUnifiedTransaction.csv).
2. Filters lines that have an order id (excluding non-order lines).
3. Groups by 'order id' and calculates the sum of 'total' (ONLY Total against Order ID).
4. Saves clean 2-column Excel report to reports/step2_payment_pivot.xlsx.
"""

from decimal import Decimal
from pathlib import Path
import sys
from typing import Optional
import pandas as pd
import xlsxwriter

from src.config import load_config
from src.payments import load_payment_report


def run_step2_payment_pivot(
    payment_file_path: Optional[str | Path] = None,
    output_excel_path: Optional[str | Path] = None,
) -> Path:
    """Execute Step 2: Pivot payment report for order id and total."""
    config = load_config()
    config.ensure_directories()

    # Locate payment file if not explicitly passed
    if payment_file_path is None:
        inbox_csv = list(config.inbox_dir.glob("*.csv"))
        raw_csv = list(config.payments_raw_dir.glob("*.csv"))
        candidates = inbox_csv or raw_csv
        if not candidates:
            raise FileNotFoundError("No payment report (.csv) found in data/inbox or data/raw/payments.")
        target_payment_file = candidates[0]
    else:
        target_payment_file = Path(payment_file_path)

    print(f"[Step 2] Reading Payment Report: {target_payment_file.name}")
    payment_res = load_payment_report(target_payment_file, config, run_id="STEP2_RUN")

    # The order-level pivot from payments: strictly Order ID and Total
    order_pivot_df = payment_res.order_pivot_df
    print(f"[Step 2] Found {len(order_pivot_df)} settled orders from {len(payment_res.lines_df)} total lines.")

    # Determine destination Excel file
    if output_excel_path is None:
        dest_path = (config.reports_dir / "step2_payment_pivot.xlsx").resolve()
    else:
        dest_path = Path(output_excel_path).resolve()
    dest_path.parent.mkdir(parents=True, exist_ok=True)

    # Write Excel with formatting and formula Grand Total
    wb = xlsxwriter.Workbook(str(dest_path))
    ws = wb.add_worksheet("Payment Pivot")

    # Formats
    fmt_title = wb.add_format({"bold": True, "font_size": 14, "font_color": "#1F4E79"})
    fmt_subtitle = wb.add_format({"italic": True, "font_size": 9, "font_color": "#595959"})
    fmt_hdr = wb.add_format({
        "bold": True, "font_size": 10, "font_color": "#FFFFFF", "bg_color": "#1F4E79",
        "border": 1, "align": "center", "valign": "vcenter"
    })
    fmt_text = wb.add_format({"border": 1, "valign": "vcenter"})
    fmt_curr = wb.add_format({"border": 1, "align": "right", "valign": "vcenter", "num_format": "₹#,##0.00"})
    fmt_curr_neg = wb.add_format({
        "bold": True, "bg_color": "#FFC7CE", "font_color": "#9C0006", "border": 1,
        "align": "right", "valign": "vcenter", "num_format": "₹#,##0.00"
    })
    fmt_total_label = wb.add_format({"bold": True, "bg_color": "#D9E1F2", "border": 1, "align": "right", "valign": "vcenter"})
    fmt_total_curr = wb.add_format({"bold": True, "bg_color": "#D9E1F2", "border": 1, "align": "right", "valign": "vcenter", "num_format": "₹#,##0.00"})

    ws.write(0, 0, "Amazon Payment Report: Order ID & Total Pivot", fmt_title)
    ws.write(1, 0, f"Source File: {target_payment_file.name} | Unique Settled Orders: {len(order_pivot_df)}", fmt_subtitle)
    ws.set_row(0, 24)
    ws.set_row(1, 16)

    headers = ["Order ID", "Total (₹)"]
    HEADER_ROW = 2
    for c_idx, h in enumerate(headers):
        ws.write(HEADER_ROW, c_idx, h, fmt_hdr)
    ws.set_row(HEADER_ROW, 26)
    ws.freeze_panes(HEADER_ROW + 1, 0)

    curr_row = HEADER_ROW + 1
    table_matrix = []

    for _, row in order_pivot_df.iterrows():
        oid = str(row.get("order_id", ""))
        tot_val = row.get("total_net", Decimal("0.00"))
        tot_float = float(tot_val) if tot_val is not None else 0.0

        table_matrix.append([oid, tot_float])

        ws.write_string(curr_row, 0, oid, fmt_text)
        ws.write_number(curr_row, 1, tot_float, fmt_curr_neg if tot_float < 0 else fmt_curr)
        curr_row += 1

    # Grand Total Row with formula
    total_row = curr_row
    ws.write(total_row, 0, "Grand Total", fmt_total_label)
    ws.write_formula(total_row, 1, f"=SUM(B4:B{total_row})", fmt_total_curr)
    ws.set_row(total_row, 20)

    # Column widths
    ws.set_column(0, 0, 28)
    ws.set_column(1, 1, 18)
    ws.autofilter(HEADER_ROW, 0, total_row - 1, len(headers) - 1)

    wb.close()
    print(f"[Step 2] Saved Step 2 Excel file to: {dest_path}")

    # Display console table
    print("\n" + "=" * 55)
    print("STEP 2: PAYMENT REPORT PIVOT (ORDER ID & TOTAL)")
    print("=" * 55)
    print(f"{'Order ID':<28} {'Total (₹)':>16}")
    print("-" * 55)
    for row in table_matrix:
        amt_str = f"₹{row[1]:,.2f}" if row[1] >= 0 else f"-₹{abs(row[1]):,.2f}"
        print(f"{row[0]:<28} {amt_str:>16}")
    total_amt = sum(r[1] for r in table_matrix)
    tot_str = f"₹{total_amt:,.2f}" if total_amt >= 0 else f"-₹{abs(total_amt):,.2f}"
    print("-" * 55)
    print(f"{'Grand Total':<28} {tot_str:>16}")
    print("=" * 55 + "\n")

    return dest_path


if __name__ == "__main__":
    file_arg = sys.argv[1] if len(sys.argv) > 1 else None
    run_step2_payment_pivot(file_arg)

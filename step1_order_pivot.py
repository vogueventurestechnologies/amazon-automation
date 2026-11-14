"""Step 1: Amazon Order Qty Pivot with SKU and Order Status Lookup.

Workflow:
1. Reads Amazon Order Report (.txt / .tsv).
2. Groups by 'amazon-order-id' and calculates sum of 'quantity' (Pivot for Qty).
3. Using 'amazon-order-id', looks up 'sku' and 'order-status'.
4. Saves clean 4-column Excel report to reports/step1_order_pivot.xlsx.
"""

from decimal import Decimal
from pathlib import Path
import sys
from typing import Optional
import pandas as pd
import xlsxwriter

from src.config import load_config
from src.master import load_master_payout, lookup_sku
from src.orders import load_order_report


def run_step1_order_pivot(
    order_file_path: Optional[str | Path] = None,
    master_file_path: Optional[str | Path] = None,
    output_excel_path: Optional[str | Path] = None,
) -> Path:
    """Execute Step 1: Pivot for quantity on amazon-order-id and lookup sku, style, payout & order-status."""
    config = load_config()
    config.ensure_directories()

    # Locate order file if not explicitly passed
    if order_file_path is None:
        inbox_txt = list(config.inbox_dir.glob("*.txt")) + list(config.inbox_dir.glob("*.tsv"))
        raw_txt = list(config.orders_raw_dir.glob("*.txt")) + list(config.orders_raw_dir.glob("*.tsv"))
        candidates = inbox_txt or raw_txt
        if not candidates:
            raise FileNotFoundError("No order report (.txt / .tsv) found in data/inbox or data/raw/orders.")
        target_order_file = candidates[0]
    else:
        target_order_file = Path(order_file_path)

    print(f"[Step 1] Reading Order Report: {target_order_file.name}")
    order_res = load_order_report(target_order_file, config, run_id="STEP1_RUN")

    summary_df = order_res.summary_df
    print(f"[Step 1] Aggregated {len(summary_df)} unique orders from {len(order_res.items_df)} item rows.")

    # Load master payout sheet (if available)
    master_map = load_master_payout(master_file_path, config=config)
    if master_map:
        print(f"[Step 1] Loaded {len(master_map)} master SKU payout records.")

    # Determine destination Excel file
    if output_excel_path is None:
        dest_path = (config.reports_dir / "step1_order_pivot.xlsx").resolve()
    else:
        dest_path = Path(output_excel_path).resolve()
    dest_path.parent.mkdir(parents=True, exist_ok=True)

    # Write Excel with formatting and formula Grand Total
    wb = xlsxwriter.Workbook(str(dest_path))
    ws = wb.add_worksheet("Order Qty Pivot")

    # Formats
    fmt_title = wb.add_format({"bold": True, "font_size": 14, "font_color": "#1F4E79"})
    fmt_subtitle = wb.add_format({"italic": True, "font_size": 9, "font_color": "#595959"})
    fmt_hdr = wb.add_format({
        "bold": True, "font_size": 10, "font_color": "#FFFFFF", "bg_color": "#1F4E79",
        "border": 1, "align": "center", "valign": "vcenter"
    })
    fmt_text = wb.add_format({"border": 1, "valign": "vcenter"})
    fmt_text_center = wb.add_format({"border": 1, "align": "center", "valign": "vcenter"})
    fmt_int = wb.add_format({"border": 1, "align": "center", "valign": "vcenter", "num_format": "#,##0"})
    fmt_curr = wb.add_format({"border": 1, "align": "right", "valign": "vcenter", "num_format": "₹#,##0.00"})
    fmt_total_label = wb.add_format({"bold": True, "bg_color": "#D9E1F2", "border": 1, "align": "right", "valign": "vcenter"})
    fmt_total_int = wb.add_format({"bold": True, "bg_color": "#D9E1F2", "border": 1, "align": "center", "valign": "vcenter", "num_format": "#,##0"})

    ws.write(0, 0, "Amazon Order Report: Quantity Pivot & SKU / Status Lookup", fmt_title)
    ws.write(1, 0, f"Source File: {target_order_file.name} | Total Orders: {len(summary_df)}", fmt_subtitle)
    ws.set_row(0, 24)
    ws.set_row(1, 16)

    headers = ["Amazon Order ID", "Quantity", "Sku Name", "STYLE NAME", "PAYOUT", "Order Status"]
    HEADER_ROW = 2
    for c_idx, h in enumerate(headers):
        ws.write(HEADER_ROW, c_idx, h, fmt_hdr)
    ws.set_row(HEADER_ROW, 26)
    ws.freeze_panes(HEADER_ROW + 1, 0)

    curr_row = HEADER_ROW + 1
    table_matrix = []

    for _, row in summary_df.iterrows():
        oid = str(row.get("amazon-order-id", ""))
        qty = int(row.get("total_qty", 0))
        sku = str(row.get("sku", ""))
        status = str(row.get("order_status", ""))

        m_info = lookup_sku(sku, master_map) if master_map else None
        style_name = m_info.style_name if m_info else ""
        payout = float(m_info.payout) if m_info and m_info.payout is not None else None

        table_matrix.append([oid, qty, sku, style_name, payout, status])

        ws.write_string(curr_row, 0, oid, fmt_text)
        ws.write_number(curr_row, 1, qty, fmt_int)
        ws.write_string(curr_row, 2, sku, fmt_text)
        ws.write_string(curr_row, 3, style_name, fmt_text)
        if payout is not None:
            ws.write_number(curr_row, 4, payout, fmt_curr)
        else:
            ws.write_string(curr_row, 4, "", fmt_text)
        ws.write_string(curr_row, 5, status, fmt_text_center)
        curr_row += 1

    # Grand Total Row with formula
    total_row = curr_row
    ws.write(total_row, 0, "Grand Total", fmt_total_label)
    ws.write_formula(total_row, 1, f"=SUM(B4:B{total_row})", fmt_total_int)
    ws.write(total_row, 2, "", fmt_total_label)
    ws.write(total_row, 3, "", fmt_total_label)
    ws.write(total_row, 4, "", fmt_total_label)
    ws.write(total_row, 5, "", fmt_total_label)
    ws.set_row(total_row, 20)

    # Column widths
    ws.set_column(0, 0, 26)
    ws.set_column(1, 1, 10)
    ws.set_column(2, 2, 26)
    ws.set_column(3, 3, 16)
    ws.set_column(4, 4, 14)
    ws.set_column(5, 5, 28)
    ws.autofilter(HEADER_ROW, 0, total_row - 1, len(headers) - 1)

    wb.close()
    print(f"[Step 1] Saved Step 1 Excel file to: {dest_path}")

    # Display console table
    print("\n" + "=" * 80)
    print("STEP 1: ORDER QUANTITY PIVOT & SKU / ORDER STATUS LOOKUP")
    print("=" * 80)
    print(f"{'Amazon Order ID':<22} {'Qty':<5} {'Sku Name':<23} {'Style Name':<12} {'Payout':<8} {'Order Status'}")
    print("-" * 90)
    for row in table_matrix:
        payout_str = f"₹{row[4]:.2f}" if row[4] is not None else "-"
        print(f"{row[0]:<22} {row[1]:<5} {row[2]:<23} {row[3]:<12} {payout_str:<8} {row[5]}")
    total_qty = sum(r[1] for r in table_matrix)
    print("-" * 90)
    print(f"{'Grand Total':<22} {total_qty:<5}")
    print("=" * 90 + "\n")

    return dest_path


if __name__ == "__main__":
    file_arg = sys.argv[1] if len(sys.argv) > 1 else None
    run_step1_order_pivot(file_arg)


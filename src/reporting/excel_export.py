"""Comprehensive 5-Sheet Excel Report Generator for Amazon.in Order + Payment Reconciliation.

Generates /reports/order_payment_recon_<from>_<to>.xlsx containing:
1. "Order Pivot": The Phase 4 table with filters, frozen header, Grand Total row.
2. "Status Summary": Count, quantity, and net amount per delivery_status, plus RVP% and RTO%.
3. "SKU Performance": Per SKU (and design/size): qty sold, RVP count, RTO count, Delivered count, return %.
4. "Non-Order Rows": Transfers, Debt pairs, Service Fees, with balance roll-forward calculation.
5. "Exceptions": Complete audit trail filterable by reason_code and severity.
"""

from datetime import datetime
from decimal import Decimal
import io
import logging
from pathlib import Path
import re
from typing import Any, Dict, List, Optional, Tuple, Union
import pandas as pd
import xlsxwriter

from src.config import AppConfig
from src.reconcile import ReconciliationResult

logger = logging.getLogger("amazon_recon.reporting")


def _to_float(val: Any) -> float:
    """Safe conversion to float for Excel numeric output."""
    if val is None or pd.isna(val):
        return 0.0
    if isinstance(val, (int, float)):
        return float(val)
    try:
        return float(Decimal(str(val)))
    except Exception:
        return 0.0


def _to_int(val: Any) -> int:
    """Safe conversion to int for Excel numeric output."""
    if val is None or pd.isna(val):
        return 0
    try:
        return int(float(str(val)))
    except Exception:
        return 0


def _extract_design_and_size(sku: str) -> Tuple[str, str]:
    """Extract design base and size suffix from apparel SKU.

    Example: 'AA(ARN)BLACK_173-M' -> ('AA(ARN)BLACK_173', 'M')
    """
    if not sku:
        return "-", "-"
    parts = sku.rsplit("-", 1)
    if len(parts) == 2 and len(parts[1]) <= 4:
        return parts[0], parts[1]
    return sku, "-"


def _build_excel_formats(workbook: xlsxwriter.Workbook) -> Dict[str, Any]:
    """Create reusable cell formatting styles for the workbook."""
    return {
        "title": workbook.add_format({
            "bold": True, "font_size": 15, "font_color": "#1F4E79", "valign": "vcenter"
        }),
        "subtitle": workbook.add_format({
            "italic": True, "font_size": 9, "font_color": "#595959", "valign": "vcenter"
        }),
        "header": workbook.add_format({
            "bold": True, "font_size": 10, "font_color": "#FFFFFF", "bg_color": "#1F4E79",
            "border": 1, "align": "center", "valign": "vcenter", "text_wrap": True
        }),
        "kpi_label": workbook.add_format({
            "bold": True, "font_size": 9, "font_color": "#333333", "bg_color": "#E9EEF4",
            "border": 1, "align": "center", "valign": "vcenter"
        }),
        "kpi_val_int": workbook.add_format({
            "bold": True, "font_size": 13, "font_color": "#1F4E79", "bg_color": "#F4F7FA",
            "border": 1, "align": "center", "valign": "vcenter", "num_format": "#,##0"
        }),
        "kpi_val_curr": workbook.add_format({
            "bold": True, "font_size": 13, "font_color": "#1F4E79", "bg_color": "#F4F7FA",
            "border": 1, "align": "center", "valign": "vcenter", "num_format": "₹#,##0.00"
        }),
        "kpi_val_pct": workbook.add_format({
            "bold": True, "font_size": 13, "font_color": "#1F4E79", "bg_color": "#F4F7FA",
            "border": 1, "align": "center", "valign": "vcenter", "num_format": "0.0%"
        }),
        "text": workbook.add_format({"border": 1, "valign": "vcenter"}),
        "text_center": workbook.add_format({"border": 1, "align": "center", "valign": "vcenter"}),
        "currency": workbook.add_format({"border": 1, "align": "right", "valign": "vcenter", "num_format": "₹#,##0.00"}),
        "integer": workbook.add_format({"border": 1, "align": "right", "valign": "vcenter", "num_format": "#,##0"}),
        "percentage": workbook.add_format({"border": 1, "align": "right", "valign": "vcenter", "num_format": "0.0%"}),
        "total_label": workbook.add_format({
            "bold": True, "bg_color": "#D9E1F2", "border": 1, "align": "right", "valign": "vcenter"
        }),
        "total_curr": workbook.add_format({
            "bold": True, "bg_color": "#D9E1F2", "border": 1, "align": "right", "valign": "vcenter",
            "num_format": "₹#,##0.00"
        }),
        "total_int": workbook.add_format({
            "bold": True, "bg_color": "#D9E1F2", "border": 1, "align": "right", "valign": "vcenter",
            "num_format": "#,##0"
        }),
        "total_pct": workbook.add_format({
            "bold": True, "bg_color": "#D9E1F2", "border": 1, "align": "right", "valign": "vcenter",
            "num_format": "0.0%"
        }),
        # Delivery status badges
        "status_delivered": workbook.add_format({
            "bold": True, "bg_color": "#C6EFCE", "font_color": "#006100", "border": 1, "align": "center"
        }),
        "status_rvp": workbook.add_format({
            "bold": True, "bg_color": "#FFC7CE", "font_color": "#9C0006", "border": 1, "align": "center"
        }),
        "status_rto": workbook.add_format({
            "bold": True, "bg_color": "#FFF2CC", "font_color": "#B25900", "border": 1, "align": "center"
        }),
        "status_cancelled": workbook.add_format({
            "bold": True, "bg_color": "#E0E0E0", "font_color": "#555555", "border": 1, "align": "center"
        }),
        "status_not_settled": workbook.add_format({
            "bold": True, "bg_color": "#FCE4D6", "font_color": "#C65911", "border": 1, "align": "center"
        }),
        "diff_negative": workbook.add_format({
            "bold": True, "bg_color": "#FFC7CE", "font_color": "#9C0006", "border": 1,
            "align": "right", "num_format": "₹#,##0.00"
        }),
        # Exception severities
        "severity_critical": workbook.add_format({
            "bold": True, "bg_color": "#C00000", "font_color": "#FFFFFF", "border": 1, "align": "center"
        }),
        "severity_error": workbook.add_format({
            "bold": True, "bg_color": "#FFC7CE", "font_color": "#9C0006", "border": 1, "align": "center"
        }),
        "severity_warning": workbook.add_format({
            "bg_color": "#FFF2CC", "font_color": "#B25900", "border": 1, "align": "center"
        }),
        "severity_info": workbook.add_format({
            "bg_color": "#DDEBF7", "font_color": "#1F4E79", "border": 1, "align": "center"
        }),
    }


def _autofit_columns(
    ws: xlsxwriter.workbook.Worksheet,
    data: List[List[Any]],
    headers: List[str],
    min_width: int = 12,
    max_width: int = 45,
) -> None:
    """Set sensible column widths based on maximum length of content."""
    for col_idx, header in enumerate(headers):
        max_len = len(str(header))
        for row in data:
            if col_idx < len(row):
                val = str(row[col_idx] if row[col_idx] is not None else "")
                max_len = max(max_len, len(val))
        calc_width = min(max(max_len + 3, min_width), max_width)
        ws.set_column(col_idx, col_idx, calc_width)


# ---------------------------------------------------------------------------
# Sheet 1: Order Pivot
# ---------------------------------------------------------------------------

def _write_order_pivot_sheet(
    workbook: xlsxwriter.Workbook,
    formats: Dict[str, Any],
    recon_df: pd.DataFrame,
    date_label: str,
    basis_used: str,
) -> None:
    """Build Sheet 1: Order Pivot table with delivery_status, filters, and Grand Total."""
    ws = workbook.add_worksheet("Order Pivot")
    ws.set_tab_color("#1F4E79")

    ws.write(0, 0, "Amazon.in Order & Payment Reconciliation Pivot", formats["title"])
    ws.write(1, 0, f"Period: {date_label} | Basis: {basis_used} | Generated: {datetime.now().strftime('%Y-%m-%d %H:%M:%S')}", formats["subtitle"])
    ws.set_row(0, 24)
    ws.set_row(1, 16)

    headers = [
        "Amazon Order ID",
        "Quantity",
        "Sku Name",
        "STYLE NAME",
        "Order Status",
        "Total Net (₹)",
        "Delivery Status",
        "Delivered Qty",
        "Return Qty",
        "NET PRODUCTION COST",
        "FINAL NET (₹)",
    ]

    HEADER_ROW = 2
    for c_idx, h in enumerate(headers):
        ws.write(HEADER_ROW, c_idx, h, formats["header"])
    ws.set_row(HEADER_ROW, 28)
    ws.freeze_panes(HEADER_ROW + 1, 0)

    current_row = HEADER_ROW + 1
    table_matrix = []

    for _, r in recon_df.iterrows():
        tot_net = r.get("total_net")
        status = r.get("delivery_status", "")
        payout = r.get("payout")
        prod_cost = r.get("net_production_cost", r.get("production_cost"))
        final_net = r.get("final_net")

        tot_net_float = _to_float(tot_net) if tot_net is not None else 0.0
        payout_float = _to_float(payout) if payout is not None and str(payout).strip() != "" else 0.0
        prod_cost_float = _to_float(prod_cost) if prod_cost is not None else (payout_float if status == "Delivered" else 0.0)
        final_net_float = _to_float(final_net) if final_net is not None else (tot_net_float - prod_cost_float)

        sku_val = str(r.get("sku_name") or r.get("sku") or "")
        style_val = str(r.get("style_name") or "")
        deliv_qty = int(r.get("delivered_qty", 0) or 0)
        ret_qty = int(r.get("return_qty", 0) or 0)

        row_vals = [
            str(r.get("amazon-order-id", "")),
            int(r.get("quantity", 0)),
            sku_val,
            style_val,
            str(r.get("order_status", "")),
            tot_net_float,
            status,
            deliv_qty,
            ret_qty,
            prod_cost_float,
            final_net_float,
        ]
        table_matrix.append(row_vals)

        ws.write_string(current_row, 0, row_vals[0], formats["text"])
        ws.write_number(current_row, 1, row_vals[1], formats["integer"])
        ws.write_string(current_row, 2, row_vals[2], formats["text"])
        ws.write_string(current_row, 3, row_vals[3], formats["text"])
        ws.write_string(current_row, 4, row_vals[4], formats["text_center"])

        # Total Net
        if tot_net_float is not None:
            net_fmt = formats["diff_negative"] if tot_net_float < 0 else formats["currency"]
            ws.write_number(current_row, 5, tot_net_float, net_fmt)
        else:
            ws.write_string(current_row, 5, "", formats["text"])

        # Delivery status badge format
        if status == "Delivered":
            st_fmt = formats["status_delivered"]
        elif status == "RVP":
            st_fmt = formats["status_rvp"]
        elif status == "RTO":
            st_fmt = formats["status_rto"]
        elif status == "Cancelled":
            st_fmt = formats["status_cancelled"]
        else:
            st_fmt = formats["status_not_settled"]

        ws.write_string(current_row, 6, status, st_fmt)

        # Delivered Qty and Return Qty
        ws.write_number(current_row, 7, deliv_qty, formats["integer"])
        ws.write_number(current_row, 8, ret_qty, formats["integer"])

        # NET PRODUCTION COST (Col 9: J)
        if status == "Delivered" and prod_cost_float > 0:
            ws.write_number(current_row, 9, prod_cost_float, formats["currency"])
        else:
            ws.write_number(current_row, 9, 0.0, formats["currency"])

        # FINAL NET (₹): Total Net - NET PRODUCTION COST (Col 10: K = F - J)
        excel_row_num = current_row + 1
        formula_final_net = f"=F{excel_row_num}-J{excel_row_num}"
        final_fmt = formats["diff_negative"] if (final_net_float is not None and final_net_float < 0) else formats["currency"]
        ws.write_formula(current_row, 10, formula_final_net, final_fmt, value=final_net_float)

        current_row += 1

    # Grand Total Row
    if table_matrix:
        total_row = current_row
        ws.write(total_row, 0, "Grand Total", formats["total_label"])

        # Qty sum (col 1: B)
        col_qty = xlsxwriter.utility.xl_col_to_name(1)
        ws.write_formula(total_row, 1, f"=SUM({col_qty}{HEADER_ROW + 2}:{col_qty}{total_row})", formats["total_int"])

        ws.write(total_row, 2, "", formats["total_label"])
        ws.write(total_row, 3, "", formats["total_label"])
        ws.write(total_row, 4, "", formats["total_label"])

        # Total Net sum (col 5: F)
        col_net = xlsxwriter.utility.xl_col_to_name(5)
        ws.write_formula(total_row, 5, f"=SUM({col_net}{HEADER_ROW + 2}:{col_net}{total_row})", formats["total_curr"])

        ws.write(total_row, 6, "", formats["total_label"])

        # Delivered Qty sum (col 7: H)
        col_deliv = xlsxwriter.utility.xl_col_to_name(7)
        ws.write_formula(total_row, 7, f"=SUM({col_deliv}{HEADER_ROW + 2}:{col_deliv}{total_row})", formats["total_int"])

        # Return Qty sum (col 8: I)
        col_ret = xlsxwriter.utility.xl_col_to_name(8)
        ws.write_formula(total_row, 8, f"=SUM({col_ret}{HEADER_ROW + 2}:{col_ret}{total_row})", formats["total_int"])

        # NET PRODUCTION COST sum (col 9: J)
        col_prod = xlsxwriter.utility.xl_col_to_name(9)
        ws.write_formula(total_row, 9, f"=SUM({col_prod}{HEADER_ROW + 2}:{col_prod}{total_row})", formats["total_curr"])

        # FINAL NET sum (col 10: K)
        col_final = xlsxwriter.utility.xl_col_to_name(10)
        ws.write_formula(total_row, 10, f"=SUM({col_final}{HEADER_ROW + 2}:{col_final}{total_row})", formats["total_curr"])

        ws.set_row(total_row, 20)

    ws.autofilter(HEADER_ROW, 0, current_row - 1 if table_matrix else HEADER_ROW, len(headers) - 1)
    _autofit_columns(ws, table_matrix, headers, min_width=12, max_width=40)


# ---------------------------------------------------------------------------
# Sheet 2: Status Summary
# ---------------------------------------------------------------------------

def _write_status_summary_sheet(
    workbook: xlsxwriter.Workbook,
    formats: Dict[str, Any],
    recon_df: pd.DataFrame,
    date_label: str,
    basis_used: str,
) -> None:
    """Build Sheet 2: Status Summary table and RVP%/RTO% metrics of shipped orders."""
    ws = workbook.add_worksheet("Status Summary")
    ws.set_tab_color("#548235")

    ws.write(0, 0, "Delivery Status Summary & Return Analysis", formats["title"])
    ws.write(1, 0, f"Period: {date_label} | Basis: {basis_used}", formats["subtitle"])
    ws.set_row(0, 24)
    ws.set_row(1, 16)

    # Compute status breakdown
    statuses = ["Delivered", "RVP", "RTO", "Not settled", "Cancelled"]
    records = []

    for st in statuses:
        sub = recon_df[recon_df["delivery_status"] == st]
        order_count = len(sub)
        total_qty = int(sub["quantity"].sum())
        net_amt = Decimal("0.00")
        for val in sub["total_net"].dropna():
            net_amt += Decimal(str(val))
        records.append({
            "status": st,
            "order_count": order_count,
            "total_qty": total_qty,
            "net_amount": _to_float(net_amt),
        })

    # Shipped orders = Delivered + RVP + RTO
    shipped_orders_count = sum(r["order_count"] for r in records if r["status"] in ("Delivered", "RVP", "RTO"))
    rvp_orders_count = next((r["order_count"] for r in records if r["status"] == "RVP"), 0)
    rto_orders_count = next((r["order_count"] for r in records if r["status"] == "RTO"), 0)
    delivered_orders_count = next((r["order_count"] for r in records if r["status"] == "Delivered"), 0)

    rvp_rate = (rvp_orders_count / shipped_orders_count) if shipped_orders_count > 0 else 0.0
    rto_rate = (rto_orders_count / shipped_orders_count) if shipped_orders_count > 0 else 0.0
    total_return_rate = ((rvp_orders_count + rto_orders_count) / shipped_orders_count) if shipped_orders_count > 0 else 0.0

    # Executive KPI Block (Rows 3-5)
    kpis = [
        ("Total Shipped Orders", shipped_orders_count, "int"),
        ("Delivered Orders", delivered_orders_count, "int"),
        ("RVP Orders", rvp_orders_count, "int"),
        ("RTO Orders", rto_orders_count, "int"),
        ("RVP Rate %", rvp_rate, "pct"),
        ("RTO Rate %", rto_rate, "pct"),
        ("Overall Return Rate %", total_return_rate, "pct"),
    ]

    for idx, (label, val, val_type) in enumerate(kpis):
        c = idx * 2
        ws.merge_range(3, c, 3, c + 1, label, formats["kpi_label"])
        fmt_key = f"kpi_val_{val_type}"
        ws.merge_range(4, c, 4, c + 1, val, formats[fmt_key])
        ws.set_column(c, c + 1, 13)

    ws.set_row(3, 18)
    ws.set_row(4, 26)

    # Status Breakdown Table
    TABLE_TITLE_ROW = 6
    ws.write(TABLE_TITLE_ROW, 0, "Delivery Status Breakdown", formats["title"])
    ws.set_row(TABLE_TITLE_ROW, 22)

    headers = [
        "Delivery Status",
        "Order Count",
        "Total Quantity",
        "Net Amount (₹)",
        "% of Shipped Orders",
        "% of Net Payout",
    ]

    HEADER_ROW = 7
    for c_idx, h in enumerate(headers):
        ws.write(HEADER_ROW, c_idx, h, formats["header"])
    ws.set_row(HEADER_ROW, 28)

    current_row = HEADER_ROW + 1
    total_net_all = sum(r["net_amount"] for r in records)

    for r in records:
        st = r["status"]
        cnt = r["order_count"]
        qty = r["total_qty"]
        amt = r["net_amount"]
        pct_shipped = (cnt / shipped_orders_count) if (shipped_orders_count > 0 and st in ("Delivered", "RVP", "RTO")) else 0.0
        pct_payout = (amt / total_net_all) if total_net_all != 0 else 0.0

        if st == "Delivered":
            st_fmt = formats["status_delivered"]
        elif st == "RVP":
            st_fmt = formats["status_rvp"]
        elif st == "RTO":
            st_fmt = formats["status_rto"]
        elif st == "Cancelled":
            st_fmt = formats["status_cancelled"]
        else:
            st_fmt = formats["status_not_settled"]

        ws.write_string(current_row, 0, st, st_fmt)
        ws.write_number(current_row, 1, cnt, formats["integer"])
        ws.write_number(current_row, 2, qty, formats["integer"])
        ws.write_number(current_row, 3, amt, formats["currency"])
        ws.write_number(current_row, 4, pct_shipped, formats["percentage"])
        ws.write_number(current_row, 5, pct_payout, formats["percentage"])
        current_row += 1

    # Grand Total Row
    total_row = current_row
    ws.write(total_row, 0, "Grand Total", formats["total_label"])
    col_cnt = xlsxwriter.utility.xl_col_to_name(1)
    ws.write_formula(total_row, 1, f"=SUM({col_cnt}{HEADER_ROW + 2}:{col_cnt}{total_row})", formats["total_int"])
    col_qty = xlsxwriter.utility.xl_col_to_name(2)
    ws.write_formula(total_row, 2, f"=SUM({col_qty}{HEADER_ROW + 2}:{col_qty}{total_row})", formats["total_int"])
    col_amt = xlsxwriter.utility.xl_col_to_name(3)
    ws.write_formula(total_row, 3, f"=SUM({col_amt}{HEADER_ROW + 2}:{col_amt}{total_row})", formats["total_curr"])
    ws.write(total_row, 4, "", formats["total_label"])
    ws.write(total_row, 5, "", formats["total_label"])
    ws.set_row(total_row, 20)

    ws.autofilter(HEADER_ROW, 0, total_row, len(headers) - 1)
    ws.set_column(0, 0, 18)
    ws.set_column(1, 2, 14)
    ws.set_column(3, 5, 20)


# ---------------------------------------------------------------------------
# Sheet 3: SKU Performance
# ---------------------------------------------------------------------------

def _write_sku_performance_sheet(
    workbook: xlsxwriter.Workbook,
    formats: Dict[str, Any],
    recon_df: pd.DataFrame,
    date_label: str,
) -> None:
    """Build Sheet 3: SKU, Design, and Size performance matrix with return rates."""
    ws = workbook.add_worksheet("SKU Performance")
    ws.set_tab_color("#7030A0")

    ws.write(0, 0, "SKU & Design Performance Matrix", formats["title"])
    ws.write(1, 0, f"Period: {date_label}", formats["subtitle"])
    ws.set_row(0, 24)
    ws.set_row(1, 16)

    headers = [
        "SKU",
        "Design Base",
        "Size",
        "Total Orders",
        "Qty Sold",
        "Delivered Count",
        "RVP Count",
        "RTO Count",
        "Cancelled Count",
        "Return Rate %",
        "Net Proceeds (₹)",
    ]

    HEADER_ROW = 2
    for c_idx, h in enumerate(headers):
        ws.write(HEADER_ROW, c_idx, h, formats["header"])
    ws.set_row(HEADER_ROW, 28)
    ws.freeze_panes(HEADER_ROW + 1, 0)

    # Group by SKU
    sku_groups: Dict[str, List[pd.Series]] = {}
    for _, row in recon_df.iterrows():
        sku = str(row["sku"] or "UNASSIGNED").strip()
        sku_groups.setdefault(sku, []).append(row)

    table_matrix = []
    current_row = HEADER_ROW + 1

    for sku, rows in sorted(sku_groups.items()):
        design, size = _extract_design_and_size(sku)
        total_orders = len(rows)
        delivered_cnt = sum(1 for r in rows if r["delivery_status"] == "Delivered")
        rvp_cnt = sum(1 for r in rows if r["delivery_status"] == "RVP")
        rto_cnt = sum(1 for r in rows if r["delivery_status"] == "RTO")
        canc_cnt = sum(1 for r in rows if r["delivery_status"] == "Cancelled")
        qty_sold = sum(int(r["quantity"]) for r in rows if r["delivery_status"] != "Cancelled")

        shipped_cnt = delivered_cnt + rvp_cnt + rto_cnt
        return_rate = ((rvp_cnt + rto_cnt) / shipped_cnt) if shipped_cnt > 0 else 0.0

        net_proceeds = sum((Decimal(str(r["total_net"])) for r in rows if r["total_net"] is not None), Decimal("0.00"))
        net_float = _to_float(net_proceeds)

        row_vals = [
            sku,
            design,
            size,
            total_orders,
            qty_sold,
            delivered_cnt,
            rvp_cnt,
            rto_cnt,
            canc_cnt,
            return_rate,
            net_float,
        ]
        table_matrix.append(row_vals)

        ws.write_string(current_row, 0, sku, formats["text"])
        ws.write_string(current_row, 1, design, formats["text"])
        ws.write_string(current_row, 2, size, formats["text_center"])
        ws.write_number(current_row, 3, total_orders, formats["integer"])
        ws.write_number(current_row, 4, qty_sold, formats["integer"])
        ws.write_number(current_row, 5, delivered_cnt, formats["integer"])
        ws.write_number(current_row, 6, rvp_cnt, formats["integer"])
        ws.write_number(current_row, 7, rto_cnt, formats["integer"])
        ws.write_number(current_row, 8, canc_cnt, formats["integer"])

        ret_fmt = formats["diff_negative"] if return_rate > 0.20 else formats["percentage"]
        ws.write_number(current_row, 9, return_rate, ret_fmt)

        p_fmt = formats["diff_negative"] if net_float < 0 else formats["currency"]
        ws.write_number(current_row, 10, net_float, p_fmt)
        current_row += 1

    # Grand Total Row
    if table_matrix:
        total_row = current_row
        ws.write(total_row, 0, "Grand Total", formats["total_label"])
        ws.write(total_row, 1, "", formats["total_label"])
        ws.write(total_row, 2, "", formats["total_label"])

        # Integer sums (cols 3 to 8)
        for col_idx in range(3, 9):
            col_letter = xlsxwriter.utility.xl_col_to_name(col_idx)
            ws.write_formula(total_row, col_idx, f"=SUM({col_letter}{HEADER_ROW + 2}:{col_letter}{total_row})", formats["total_int"])

        # Overall Return Rate: (Total RVP + Total RTO) / (Total Delivered + RVP + RTO)
        col_del = xlsxwriter.utility.xl_col_to_name(5)
        col_rvp = xlsxwriter.utility.xl_col_to_name(6)
        col_rto = xlsxwriter.utility.xl_col_to_name(7)
        ws.write_formula(
            total_row, 9,
            f"=IF(({col_del}{total_row + 1}+{col_rvp}{total_row + 1}+{col_rto}{total_row + 1})>0, "
            f"({col_rvp}{total_row + 1}+{col_rto}{total_row + 1})/({col_del}{total_row + 1}+{col_rvp}{total_row + 1}+{col_rto}{total_row + 1}), 0)",
            formats["total_pct"]
        )

        # Net proceeds sum
        col_net = xlsxwriter.utility.xl_col_to_name(10)
        ws.write_formula(total_row, 10, f"=SUM({col_net}{HEADER_ROW + 2}:{col_net}{total_row})", formats["total_curr"])
        ws.set_row(total_row, 20)

    ws.autofilter(HEADER_ROW, 0, current_row - 1 if table_matrix else HEADER_ROW, len(headers) - 1)
    _autofit_columns(ws, table_matrix, headers, min_width=12, max_width=35)


# ---------------------------------------------------------------------------
# Sheet 4: Non-Order Rows
# ---------------------------------------------------------------------------

def _write_non_order_rows_sheet(
    workbook: xlsxwriter.Workbook,
    formats: Dict[str, Any],
    non_order_df: pd.DataFrame,
    lines_df: pd.DataFrame,
    grand_total: Decimal,
    date_label: str,
) -> None:
    """Build Sheet 4: Transfers, Debt pairs, Service Fees and Balance Roll-Forward."""
    ws = workbook.add_worksheet("Non-Order Rows")
    ws.set_tab_color("#FFC000")

    ws.write(0, 0, "Non-Order Financial Rows & Amazon Balance Roll-Forward", formats["title"])
    ws.write(1, 0, f"Period: {date_label}", formats["subtitle"])
    ws.set_row(0, 24)
    ws.set_row(1, 16)

    # Filter non-order transaction lines from lines_df
    non_order_lines = lines_df[lines_df["clean_order_id"] == ""].copy()

    # Roll-forward summary block (Rows 3-6)
    transfers_total = non_order_lines[non_order_lines["type"].str.lower() == "transfer"]["total"].sum() if not non_order_lines.empty else Decimal("0.00")
    debt_total = non_order_lines[non_order_lines["type"].str.lower() == "debt"]["total"].sum() if not non_order_lines.empty else Decimal("0.00")
    service_fees_total = non_order_lines[non_order_lines["type"].str.lower() == "service fee"]["total"].sum() if not non_order_lines.empty else Decimal("0.00")
    order_lines_total = lines_df[lines_df["clean_order_id"] != ""]["total"].sum() if not lines_df.empty else Decimal("0.00")

    # Closing Amazon Balance Roll-forward = Grand Total - Transfers (since transfers represent funds withdrawn)
    closing_balance = grand_total - transfers_total

    kpi_defs = [
        ("Gross Line Total", _to_float(grand_total), "curr"),
        ("Payout Transfers (Bank)", _to_float(transfers_total), "curr"),
        ("Debt Pairs Net", _to_float(debt_total), "curr"),
        ("Service Fees", _to_float(service_fees_total), "curr"),
        ("Order Lines Net", _to_float(order_lines_total), "curr"),
        ("Closing Balance Roll-Forward", _to_float(closing_balance), "curr"),
    ]

    for idx, (label, val, val_type) in enumerate(kpi_defs):
        c = idx * 2
        ws.merge_range(3, c, 3, c + 1, label, formats["kpi_label"])
        fmt_key = f"kpi_val_{val_type}"
        ws.merge_range(4, c, 4, c + 1, val, formats[fmt_key])
        ws.set_column(c, c + 1, 13)

    ws.set_row(3, 18)
    ws.set_row(4, 26)

    # Detailed non-order transactions table
    TABLE_TITLE_ROW = 6
    ws.write(TABLE_TITLE_ROW, 0, "Non-Order Transaction Details (Transfers, Debt, Service Fees)", formats["title"])
    ws.set_row(TABLE_TITLE_ROW, 22)

    headers = [
        "Settlement ID",
        "Date (IST)",
        "Transaction Type",
        "Description",
        "Amount (₹)",
    ]

    HEADER_ROW = 7
    for c_idx, h in enumerate(headers):
        ws.write(HEADER_ROW, c_idx, h, formats["header"])
    ws.set_row(HEADER_ROW, 28)
    ws.freeze_panes(HEADER_ROW + 1, 0)

    current_row = HEADER_ROW + 1
    table_matrix = []

    for _, r in non_order_lines.iterrows():
        amt_float = _to_float(r["total"])
        row_vals = [
            r["settlement_id"],
            r["tx_date_ist"] or "-",
            r["type"],
            r["description"] or "-",
            amt_float,
        ]
        table_matrix.append(row_vals)

        ws.write_string(current_row, 0, str(row_vals[0]), formats["text"])
        ws.write_string(current_row, 1, str(row_vals[1]), formats["text_center"])
        ws.write_string(current_row, 2, str(row_vals[2]), formats["text_center"])
        ws.write_string(current_row, 3, str(row_vals[3]), formats["text"])

        amt_fmt = formats["diff_negative"] if amt_float < 0 else formats["currency"]
        ws.write_number(current_row, 4, amt_float, amt_fmt)
        current_row += 1

    # Grand Total Row
    if table_matrix:
        total_row = current_row
        ws.write(total_row, 0, "Total Non-Order Net", formats["total_label"])
        ws.write(total_row, 1, "", formats["total_label"])
        ws.write(total_row, 2, "", formats["total_label"])
        ws.write(total_row, 3, "", formats["total_label"])
        col_amt = xlsxwriter.utility.xl_col_to_name(4)
        ws.write_formula(total_row, 4, f"=SUM({col_amt}{HEADER_ROW + 2}:{col_amt}{total_row})", formats["total_curr"])
        ws.set_row(total_row, 20)

    ws.autofilter(HEADER_ROW, 0, current_row - 1 if table_matrix else HEADER_ROW, len(headers) - 1)
    _autofit_columns(ws, table_matrix, headers, min_width=14, max_width=45)


# ---------------------------------------------------------------------------
# Sheet 5: Exceptions
# ---------------------------------------------------------------------------

def _write_exceptions_sheet(
    workbook: xlsxwriter.Workbook,
    formats: Dict[str, Any],
    exceptions: List[Any],
    date_label: str,
) -> None:
    """Build Sheet 5: Complete audit trail filterable by reason_code and severity."""
    ws = workbook.add_worksheet("Exceptions")
    ws.set_tab_color("#C00000")

    ws.write(0, 0, "Reconciliation Exceptions & Data Quality Audit Trail", formats["title"])
    ws.write(1, 0, f"Period: {date_label}", formats["subtitle"])
    ws.set_row(0, 24)
    ws.set_row(1, 16)

    headers = [
        "Run ID",
        "Source Component",
        "Reference ID",
        "Reason Code",
        "Severity",
        "Details",
    ]

    HEADER_ROW = 2
    for c_idx, h in enumerate(headers):
        ws.write(HEADER_ROW, c_idx, h, formats["header"])
    ws.set_row(HEADER_ROW, 28)
    ws.freeze_panes(HEADER_ROW + 1, 0)

    current_row = HEADER_ROW + 1
    table_matrix = []

    for exc in exceptions:
        sev = str(exc.severity).upper()
        row_vals = [
            exc.run_id,
            exc.source,
            exc.reference_id or "-",
            exc.reason_code,
            sev,
            exc.details,
        ]
        table_matrix.append(row_vals)

        ws.write_string(current_row, 0, str(row_vals[0]), formats["text"])
        ws.write_string(current_row, 1, str(row_vals[1]), formats["text"])
        ws.write_string(current_row, 2, str(row_vals[2]), formats["text"])
        ws.write_string(current_row, 3, str(row_vals[3]), formats["text"])

        # Severity style
        if sev == "CRITICAL":
            s_fmt = formats["severity_critical"]
        elif sev == "ERROR":
            s_fmt = formats["severity_error"]
        elif sev == "WARNING":
            s_fmt = formats["severity_warning"]
        else:
            s_fmt = formats["severity_info"]

        ws.write_string(current_row, 4, sev, s_fmt)
        ws.write_string(current_row, 5, str(row_vals[5]), formats["text"])
        current_row += 1

    ws.autofilter(HEADER_ROW, 0, current_row - 1 if table_matrix else HEADER_ROW, len(headers) - 1)
    _autofit_columns(ws, table_matrix, headers, min_width=14, max_width=50)


def _write_payment_only_sheet(
    workbook: xlsxwriter.Workbook,
    formats: Dict[str, Any],
    payment_only_df: pd.DataFrame,
    date_label: str,
) -> None:
    """Build Sheet: Payment Only Orders (excluded from main master sheet)."""
    ws = workbook.add_worksheet("Payment Only Orders")
    ws.set_tab_color("#7F7F7F")

    ws.write(0, 0, "Payment Orders Not Found in Order Report (Excluded from Master Reconciliation)", formats["title"])
    ws.write(1, 0, f"Period: {date_label} | Total Excluded: {len(payment_only_df)}", formats["subtitle"])
    ws.set_row(0, 24)
    ws.set_row(1, 16)

    headers = [
        "Amazon Order ID",
        "SKU",
        "Quantity",
        "Total Net (₹)",
        "Net Ex Shipping (₹)",
        "Settlement IDs",
    ]

    HEADER_ROW = 2
    for c_idx, h in enumerate(headers):
        ws.write(HEADER_ROW, c_idx, h, formats["header"])
    ws.set_row(HEADER_ROW, 28)
    ws.freeze_panes(HEADER_ROW + 1, 0)

    current_row = HEADER_ROW + 1
    table_matrix = []

    for _, r in payment_only_df.iterrows():
        tot_net_f = _to_float(r.get("total_net", 0.0))
        net_ex_f = _to_float(r.get("net_ex_shipping", 0.0))
        row_vals = [
            str(r.get("order_id", "")),
            str(r.get("sku", "")),
            _to_int(r.get("quantity", 0)),
            tot_net_f,
            net_ex_f,
            str(r.get("settlement_ids", "")),
        ]
        table_matrix.append(row_vals)

        ws.write_string(current_row, 0, row_vals[0], formats["text"])
        ws.write_string(current_row, 1, row_vals[1], formats["text"])
        ws.write_number(current_row, 2, row_vals[2], formats["integer"])
        ws.write_number(current_row, 3, row_vals[3], formats["diff_negative"] if tot_net_f < 0 else formats["currency"])
        ws.write_number(current_row, 4, row_vals[4], formats["diff_negative"] if net_ex_f < 0 else formats["currency"])
        ws.write_string(current_row, 5, row_vals[5], formats["text"])
        current_row += 1

    ws.autofilter(HEADER_ROW, 0, current_row - 1 if table_matrix else HEADER_ROW, len(headers) - 1)
    _autofit_columns(ws, table_matrix, headers, min_width=14, max_width=35)


# ---------------------------------------------------------------------------
# Main Excel Report Generation Function
# ---------------------------------------------------------------------------

def generate_reconciliation_excel_report(
    recon_res: ReconciliationResult,
    config: AppConfig,
    lines_df: Optional[pd.DataFrame] = None,
    from_date: Optional[str] = None,
    to_date: Optional[str] = None,
    output_path: Optional[Union[Path, str, io.BytesIO]] = None,
    save_latest_copy: bool = True,
    ephemeral: bool = False,
) -> Union[Path, io.BytesIO]:
    """Generate the complete 5-sheet reconciliation Excel workbook.

    Args:
        recon_res: ReconciliationResult from Phase 4.
        config: Application configuration.
        lines_df: Raw parsed payment lines from Phase 3.
        from_date: Optional start date filter (IST, YYYY-MM-DD).
        to_date: Optional end date filter (IST, YYYY-MM-DD).
        output_path: Optional destination path or io.BytesIO buffer.
        save_latest_copy: Whether to copy to order_payment_recon_latest.xlsx (default True).
        ephemeral: If True, generates workbook purely in-memory (io.BytesIO) and writes nothing to disk.

    Returns:
        Path to the generated Excel file, or io.BytesIO in ephemeral mode.
    """
    # Determine date range label
    if from_date and to_date:
        date_label = f"{from_date} to {to_date}"
        file_suffix = f"{from_date}_{to_date}"
    elif from_date:
        date_label = f"From {from_date} onwards"
        file_suffix = f"from_{from_date}"
    elif to_date:
        date_label = f"Up to {to_date}"
        file_suffix = f"to_{to_date}"
    else:
        date_label = "All available dates"
        file_suffix = "all"

    # Ephemeral / in-memory mode: zero disk footprint
    if ephemeral or isinstance(output_path, io.BytesIO):
        buffer = output_path if isinstance(output_path, io.BytesIO) else io.BytesIO()
        workbook = xlsxwriter.Workbook(buffer, {"strings_to_numbers": False, "in_memory": True})
        formats = _build_excel_formats(workbook)

        # 1. Sheet 1: Order Pivot
        _write_order_pivot_sheet(workbook, formats, recon_res.recon_df, date_label, recon_res.basis_used)

        # 2. Optional Sheet: Payment Only Orders (only if configured)
        if getattr(config, "include_payment_only_orders", False) and not recon_res.payment_only_df.empty:
            _write_payment_only_sheet(workbook, formats, recon_res.payment_only_df, date_label)

        workbook.close()
        buffer.seek(0)
        logger.info("Excel workbook generated in-memory (ephemeral mode).")
        return buffer

    config.ensure_directories()

    # Resolve output path
    if output_path is None:
        target_file = (config.reports_dir / f"order_payment_recon_{file_suffix}.xlsx").resolve()
    else:
        target_file = Path(output_path).resolve()
    target_file.parent.mkdir(parents=True, exist_ok=True)

    logger.info(f"Writing Excel reconciliation workbook to: {target_file}")

    workbook = xlsxwriter.Workbook(str(target_file), {"strings_to_numbers": False})
    formats = _build_excel_formats(workbook)

    # 1. Sheet 1: Order Pivot
    _write_order_pivot_sheet(workbook, formats, recon_res.recon_df, date_label, recon_res.basis_used)

    # 2. Optional Sheet: Payment Only Orders (only if configured)
    if getattr(config, "include_payment_only_orders", False) and not recon_res.payment_only_df.empty:
        _write_payment_only_sheet(workbook, formats, recon_res.payment_only_df, date_label)

    workbook.close()
    logger.info(f"Excel workbook generated successfully: {target_file}")

    # Also update convenience alias order_payment_recon_latest.xlsx if enabled
    if save_latest_copy:
        latest_file = config.reports_dir / "order_payment_recon_latest.xlsx"
        try:
            import shutil
            shutil.copy2(target_file, latest_file)
            logger.info(f"Updated latest report copy: {latest_file}")
        except Exception as e:
            logger.debug(f"Could not copy latest report: {e}")

    return target_file

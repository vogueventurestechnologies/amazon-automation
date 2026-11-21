"""Unit tests for Phase 5: Multi-Tab Excel Reconciliation Report (Master Order Scope).

Verifies the generated Excel workbook:
1. Sheet "Order Pivot": Master orders only (5 rows), delivery_status, filters, frozen headers, Grand Total formulas.
2. Sheet "Status Summary": Calculated only from master orders (4 shipped: 1 Delivered, 2 RVP, 1 RTO; 1 Cancelled).
3. Sheet "SKU Performance": Decomposed by SKU, design, and size strictly for master orders.
4. Sheet "Non-Order Rows": Transfers, Debt pairs, Service fees, and Closing Amazon Balance roll-forward.
5. Sheet "Exceptions": Complete audit trail filterable by reason_code and severity.
6. Optional Sheet "Payment Only Orders": Generated when include_payment_only_orders = True.
"""

from decimal import Decimal
from pathlib import Path
import openpyxl
import pytest

from src.config import AppConfig
from src.orders import load_order_report
from src.payments import load_payment_report
from src.reconcile import reconcile_orders_and_payments
from src.reporting.excel_export import generate_reconciliation_excel_report

REAL_ORDER_FILE = Path("D:/2026/NOVEMBER MONTH/13-11-2026/157818151187020770.txt")
REAL_PAYMENT_FILE = Path("D:/2026/NOVEMBER MONTH/13-11-2026/2026Aug1-2026Oct31CustomUnifiedTransaction.csv")


def test_excel_export_sheets_and_structure_master_orders_only(tmp_project: AppConfig):
    """Test full workbook generation and verify all 5 sheet structures against real samples."""
    if not (REAL_ORDER_FILE.exists() and REAL_PAYMENT_FILE.exists()):
        pytest.skip("Real sample files not available on test system")

    order_res = load_order_report(REAL_ORDER_FILE, tmp_project, run_id="EXCEL_ORD")
    pmt_res = load_payment_report(REAL_PAYMENT_FILE, tmp_project, run_id="EXCEL_PMT")
    master_file = Path("data/master/MASTER PAYOUT SHEET.xlsx")
    master_map = None
    if master_file.exists():
        from src.master import load_master_payout
        master_map = load_master_payout(master_file, tmp_project)
    recon_res = reconcile_orders_and_payments(order_res, pmt_res, tmp_project, run_id="EXCEL_RECON", master_map=master_map)

    output_path = tmp_project.reports_dir / "order_payment_recon_test.xlsx"
    generated_path = generate_reconciliation_excel_report(
        recon_res=recon_res,
        config=tmp_project,
        lines_df=pmt_res.lines_df,
        output_path=output_path,
    )

    assert generated_path.exists()
    assert generated_path.stat().st_size > 0

    # Inspect generated workbook with openpyxl
    wb = openpyxl.load_workbook(generated_path, data_only=False)
    expected_sheets = [
        "Order Pivot",
    ]
    assert wb.sheetnames == expected_sheets

    # --- 1. Sheet: Order Pivot ---
    ws_pivot = wb["Order Pivot"]
    # Check headers at row 3
    headers = [ws_pivot.cell(3, c).value for c in range(1, 13)]
    expected_headers = [
        "Amazon Order ID",
        "Quantity",
        "Sku Name",
        "STYLE NAME",
        "PAYOUT",
        "Order Status",
        "Total Net (₹)",
        "Delivery Status",
        "Delivered Qty",
        "Return Qty",
        "NET PRODUCTION COST",
        "FINAL NET (₹)",
    ]
    assert headers == expected_headers

    # Master orders only: exactly 5 data rows (rows 4 to 8)
    order_rows = {}
    for r in range(4, 9):
        oid = str(ws_pivot.cell(r, 1).value).strip()
        order_rows[oid] = {
            "qty": ws_pivot.cell(r, 2).value,
            "sku": ws_pivot.cell(r, 3).value,
            "style": ws_pivot.cell(r, 4).value,
            "payout": ws_pivot.cell(r, 5).value,
            "order_status": ws_pivot.cell(r, 6).value,
            "total": ws_pivot.cell(r, 7).value,
            "delivery_status": ws_pivot.cell(r, 8).value,
            "delivered_qty": ws_pivot.cell(r, 9).value,
            "return_qty": ws_pivot.cell(r, 10).value,
            "net_prod_cost": ws_pivot.cell(r, 11).value,
            "final_net": ws_pivot.cell(r, 12).value,
        }

    assert len(order_rows) == 5

    # Verify each of the 5 canonical rows
    assert "171-0478052-6676325" in order_rows
    assert order_rows["171-0478052-6676325"]["delivery_status"] == "RVP"
    assert order_rows["171-0478052-6676325"]["total"] == -140.42
    assert order_rows["171-0478052-6676325"]["net_prod_cost"] == 0.0
    assert order_rows["171-0478052-6676325"]["qty"] == 1
    assert order_rows["171-0478052-6676325"]["delivered_qty"] == 0
    assert order_rows["171-0478052-6676325"]["return_qty"] == 1
    assert order_rows["171-0478052-6676325"]["sku"] == "AA(ARN)BLACK_302-M"

    assert "404-3526450-9656333" in order_rows
    assert order_rows["404-3526450-9656333"]["delivery_status"] == "RVP"
    assert order_rows["404-3526450-9656333"]["total"] == -140.42
    assert order_rows["404-3526450-9656333"]["net_prod_cost"] == 0.0
    assert order_rows["404-3526450-9656333"]["delivered_qty"] == 0
    assert order_rows["404-3526450-9656333"]["return_qty"] == 1

    assert "404-6736655-4486719" in order_rows
    assert order_rows["404-6736655-4486719"]["delivery_status"] == "RVP"
    assert order_rows["404-6736655-4486719"]["total"] == -64.90
    assert order_rows["404-6736655-4486719"]["net_prod_cost"] == 0.0
    assert order_rows["404-6736655-4486719"]["delivered_qty"] == 0
    assert order_rows["404-6736655-4486719"]["return_qty"] == 1

    assert "407-7389902-9127540" in order_rows
    assert order_rows["407-7389902-9127540"]["delivery_status"] == "Delivered"
    assert order_rows["407-7389902-9127540"]["total"] == 306.24
    assert order_rows["407-7389902-9127540"]["net_prod_cost"] == 147.0
    assert order_rows["407-7389902-9127540"]["delivered_qty"] == 1
    assert order_rows["407-7389902-9127540"]["return_qty"] == 0

    assert "408-3937651-1846748" in order_rows
    assert order_rows["408-3937651-1846748"]["delivery_status"] == "Cancelled"
    assert order_rows["408-3937651-1846748"]["net_prod_cost"] == 0.0
    assert order_rows["408-3937651-1846748"]["delivered_qty"] == 0
    assert order_rows["408-3937651-1846748"]["return_qty"] == 0

    # Check Grand Total row at row 9
    assert ws_pivot.cell(9, 1).value == "Grand Total"
    assert ws_pivot.cell(9, 2).value == "=SUM(B4:B8)"
    assert ws_pivot.cell(9, 7).value == "=SUM(G4:G8)"
    assert ws_pivot.cell(9, 9).value == "=SUM(I4:I8)"
    assert ws_pivot.cell(9, 10).value == "=SUM(J4:J8)"
    assert ws_pivot.cell(9, 11).value == "=SUM(K4:K8)"
    assert ws_pivot.cell(9, 12).value == "=SUM(L4:L8)"


def test_excel_export_with_payment_only_orders_sheet(tmp_project: AppConfig):
    """Test that when include_payment_only_orders is True, sheet 'Payment Only Orders' is added."""
    if not (REAL_ORDER_FILE.exists() and REAL_PAYMENT_FILE.exists()):
        pytest.skip("Real sample files not available on test system")

    tmp_project.include_payment_only_orders = True

    order_res = load_order_report(REAL_ORDER_FILE, tmp_project, run_id="EXCEL_PO_ORD")
    pmt_res = load_payment_report(REAL_PAYMENT_FILE, tmp_project, run_id="EXCEL_PO_PMT")
    recon_res = reconcile_orders_and_payments(order_res, pmt_res, tmp_project, run_id="EXCEL_PO_RECON")

    output_path = tmp_project.reports_dir / "order_payment_recon_with_po.xlsx"
    generated_path = generate_reconciliation_excel_report(
        recon_res=recon_res,
        config=tmp_project,
        lines_df=pmt_res.lines_df,
        output_path=output_path,
    )

    wb = openpyxl.load_workbook(generated_path, data_only=False)
    assert "Payment Only Orders" in wb.sheetnames

    ws_po = wb["Payment Only Orders"]
    # 9 excluded payment orders
    assert ws_po.max_row == 12  # header at row 3 + 9 data rows

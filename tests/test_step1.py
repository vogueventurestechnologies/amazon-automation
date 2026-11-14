"""Unit tests for Step 1: Order Qty Pivot and SKU/Status Lookup."""

from pathlib import Path
import openpyxl
import pytest

from step1_order_pivot import run_step1_order_pivot

ORDER_FILE = Path("data/raw/orders/157818151187020770.txt")


def test_step1_order_pivot_execution(tmp_path: Path):
    """Test Step 1 script generates expected 4-column workbook with grand total formula."""
    if not ORDER_FILE.exists():
        pytest.skip("Order file not found")

    out_file = tmp_path / "step1_test.xlsx"
    res_path = run_step1_order_pivot(ORDER_FILE, out_file)

    assert res_path.exists()
    wb = openpyxl.load_workbook(res_path, data_only=False)
    assert "Order Qty Pivot" in wb.sheetnames

    ws = wb["Order Qty Pivot"]
    # Check headers at row 3
    headers = [ws.cell(3, c).value for c in range(1, 7)]
    assert headers == ["Amazon Order ID", "Quantity", "Sku Name", "STYLE NAME", "PAYOUT", "Order Status"]

    # Exactly 5 orders (rows 4 to 8)
    rows = {}
    for r in range(4, 9):
        oid = ws.cell(r, 1).value
        rows[oid] = {
            "qty": ws.cell(r, 2).value,
            "sku": ws.cell(r, 3).value,
            "style": ws.cell(r, 4).value,
            "payout": ws.cell(r, 5).value,
            "status": ws.cell(r, 6).value,
        }

    assert len(rows) == 5
    assert "404-6736655-4486719" in rows
    assert rows["404-6736655-4486719"]["qty"] == 1
    assert rows["404-6736655-4486719"]["sku"] == "AA(ARN)BLACK_173-L"
    assert rows["404-6736655-4486719"]["style"] == "ARN_BLACK"
    assert rows["404-6736655-4486719"]["payout"] == 156
    assert rows["404-6736655-4486719"]["status"] == "Shipped - Delivered to Buyer"

    assert "408-3937651-1846748" in rows
    assert rows["408-3937651-1846748"]["qty"] == 0
    assert rows["408-3937651-1846748"]["status"] == "Cancelled"

    # Check Grand Total row at row 9
    assert ws.cell(9, 1).value == "Grand Total"
    assert ws.cell(9, 2).value == "=SUM(B4:B8)"


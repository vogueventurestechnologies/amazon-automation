"""Unit tests for Step 2: Payment Report Pivot on Order ID and Total."""

from pathlib import Path
import openpyxl
import pytest

from step2_payment_pivot import run_step2_payment_pivot

PMT_FILE = Path("data/raw/payments/2026Aug1-2026Oct31CustomUnifiedTransaction.csv")


def test_step2_payment_pivot_execution(tmp_path: Path):
    """Test Step 2 generates clean 2-column Excel workbook with grand total formula."""
    if not PMT_FILE.exists():
        pytest.skip("Payment file not found")

    out_file = tmp_path / "step2_test.xlsx"
    res_path = run_step2_payment_pivot(PMT_FILE, out_file)

    assert res_path.exists()
    wb = openpyxl.load_workbook(res_path, data_only=False)
    assert "Payment Pivot" in wb.sheetnames

    ws = wb["Payment Pivot"]
    # Check headers at row 3
    headers = [ws.cell(3, 1).value, ws.cell(3, 2).value]
    assert headers == ["Order ID", "Total (₹)"]

    # 13 settled orders (rows 4 to 16)
    rows = {}
    for r in range(4, 17):
        oid = ws.cell(r, 1).value
        rows[oid] = ws.cell(r, 2).value

    assert len(rows) == 13
    assert "407-7389902-9127540" in rows
    assert rows["407-7389902-9127540"] == 306.24
    assert rows["171-0478052-6676325"] == -140.42
    assert rows["404-6736655-4486719"] == -64.90

    # Grand Total at row 17
    assert ws.cell(17, 1).value == "Grand Total"
    assert ws.cell(17, 2).value == "=SUM(B4:B16)"

"""Unit tests for Phase 3: Load Payment Report, Debt Validation, and Order-Level Pivot."""

from decimal import Decimal
from pathlib import Path
import pandas as pd
import pytest

from src.config import AppConfig
from src.payments import (
    find_header_row_index,
    is_payment_report,
    load_payment_report,
    normalize_and_parse_date,
)

REAL_PAYMENT_FILE = Path("D:/2026/NOVEMBER MONTH/13-11-2026/2026Aug1-2026Oct31CustomUnifiedTransaction.csv")


# Canonical sample fixture with preamble, definitions, orders, refunds, shipping, debts, and transfers
SAMPLE_PREAMBLE_AND_HEADER = (
    '"Includes Amazon Marketplace, Fulfillment by Amazon (FBA) and Amazon Payments transactions"\n'
    '"All amounts in INR, unless specified"\n'
    '"Definitions:"\n'
    '"Date/Time: Posted date/time of the transaction"\n'
    '"Total sales tax liable:Includes GST or VAT liability."\n'
    '"TCS-CGST:Tax Collected at Source from CGST."\n'
    '"TCS-SGST:Tax Collected at Source from SGST."\n'
    '"TCS-IGST:Tax Collected at Source from IGST."\n'
    '"TDS (Section 194-O): Tax Deducted at Source."\n'
    '"Selling fees: Includes variable closing fee, referral fee."\n'
    '"Other transaction fees: Includes shipping, weight handling fee."\n'
    '"Other: Includes non-order transaction amount."\n'
    '"Note: Reports with start date of 1 Jan 2026 onwards."\n'
    '"date/time","settlement id","type","order id","Sku","description","quantity","marketplace","account type","fulfillment","order city","order state","order postal","product sales","shipping credits","gift wrap credits","promotional rebates","Total sales tax liable(GST before adjusting TCS)","TCS-CGST","TCS-SGST","TCS-IGST","TDS (Section 194-O)","selling fees","fba fees","other transaction fees","other","total","Transaction Status","Transaction Release Date"\n'
)

SAMPLE_ROWS = (
    # 1. Order 407-7389902-9127540 (Order +371.14, Shipping -64.90 -> total_net 306.24)
    '"10 Aug 2026 8:17:01 am UTC","27652530972","Order","407-7389902-9127540","AA(ARN)BLACK_43-M","T-Shirt","1","amazon.in","Standard","Merchant","DHANBAD","JHARKHAND","826004","380.00","0.00","0.00","0.00","19.00","-1.90","-1.90","0.00","-3.80","-20.26","0.00","0.00","0.00","371.14","Released","15 Aug 2026"\n'
    '"11 Aug 2026 8:17:01 am UTC","27652530972","Shipping Services","407-7389902-9127540","","Easy Ship Fee","","amazon.in","Standard","Merchant","","","","0.00","0.00","0.00","0.00","0.00","0.00","0.00","0.00","0.00","0.00","0.00","-64.90","0.00","-64.90","Released","15 Aug 2026"\n'

    # 2. Order 171-0478052-6676325 (Order +371.14, Shipping -64.90, Refund -446.66 -> total_net -140.42, net_ex_shipping -75.52)
    '"12 Aug 2026 8:17:01 am UTC","27652530972","Order","171-0478052-6676325","AA(ARN)BLACK_302-M","T-Shirt","1","amazon.in","Standard","Merchant","AURANGABAD","MAHARASHTRA","431005","380.00","0.00","0.00","0.00","19.00","-1.90","-1.90","0.00","-3.80","-20.26","0.00","0.00","0.00","371.14","Released","15 Aug 2026"\n'
    '"13 Aug 2026 8:17:01 am UTC","27652530972","Shipping Services","171-0478052-6676325","","Easy Ship Fee","","amazon.in","Standard","Merchant","","","","0.00","0.00","0.00","0.00","0.00","0.00","0.00","0.00","0.00","0.00","0.00","-64.90","0.00","-64.90","Released","15 Aug 2026"\n'
    '"14 Aug 2026 8:17:01 am UTC","27652530972","Refund","171-0478052-6676325","AA(ARN)BLACK_302-M","Refund","1","amazon.in","Standard","Merchant","AURANGABAD","MAHARASHTRA","431005","-380.00","0.00","0.00","0.00","-19.00","1.90","1.90","0.00","3.80","2.00","0.00","0.00","-57.26","-446.66","Released","15 Aug 2026"\n'

    # 3. Order 404-6736655-4486719 (Order 0.00, Shipping -64.90 -> total_net -64.90, net_ex_shipping 0.00)
    '"22 Aug 2026 8:17:01 am UTC","27652530972","Order","404-6736655-4486719","AA(ARN)BLACK_173-L","Replacement","1","amazon.in","Standard","Merchant","SAGAR","KARNATAKA","577401","0.00","0.00","0.00","0.00","0.00","0.00","0.00","0.00","0.00","0.00","0.00","0.00","0.00","0.00","Released","28 Aug 2026"\n'
    '"23 Aug 2026 8:17:01 am UTC","27788990011","Shipping Services","404-6736655-4486719","","Easy Ship Fee","","amazon.in","Standard","Merchant","","","","0.00","0.00","0.00","0.00","0.00","0.00","0.00","0.00","0.00","0.00","0.00","-64.90","0.00","-64.90","Released","28 Aug 2026"\n'

    # 4. Debt Pair (nets to 0)
    '"25 Aug 2026 8:17:01 am UTC","27652530972","Debt","","","Cross Account Transfer","","amazon.in","Standard","Merchant","","","","0.00","0.00","0.00","0.00","0.00","0.00","0.00","0.00","0.00","0.00","0.00","0.00","-50.00","-50.00","Released","28 Aug 2026"\n'
    '"25 Aug 2026 8:17:01 am UTC","27652530972","Debt","","","Cross Account Transfer","","amazon.in","Standard","Merchant","","","","0.00","0.00","0.00","0.00","0.00","0.00","0.00","0.00","0.00","0.00","0.00","0.00","50.00","50.00","Released","28 Aug 2026"\n'

    # 5. Non-order Transfer row (payout)
    '"26 Aug 2026 8:17:01 am UTC","27652530972","Transfer","","","To account ending with: 863","","amazon.in","Standard","Merchant","","","","0.00","0.00","0.00","0.00","0.00","0.00","0.00","0.00","0.00","0.00","0.00","0.00","0.00","-101.44","Released","28 Aug 2026"\n'
)


@pytest.fixture
def sample_payment_file(tmp_path: Path) -> Path:
    """Provide a self-contained sample payment CSV with 13 preamble lines."""
    f = tmp_path / "sample_payment.csv"
    f.write_text(SAMPLE_PREAMBLE_AND_HEADER + SAMPLE_ROWS, encoding="utf-8-sig")
    return f


def test_is_payment_report_detection(sample_payment_file: Path, tmp_path: Path):
    """Verify is_payment_report detects Custom Unified Transaction payment reports."""
    assert is_payment_report(sample_payment_file) is True

    non_payment = tmp_path / "order.txt"
    non_payment.write_text("amazon-order-id\torder-item-id\tsku\n", encoding="utf-8")
    assert is_payment_report(non_payment) is False


def test_dynamic_header_detection_skiprows(sample_payment_file: Path):
    """Verify find_header_row_index dynamically skips the 13 preamble lines and finds line 13."""
    header_idx = find_header_row_index(sample_payment_file)
    assert header_idx == 13


def test_sept_date_normalization():
    """Verify '3 Sept 2026 8:17:04 am UTC' with 'Sept' normalizes and parses cleanly."""
    date_str = "3 Sept 2026 8:17:04 am UTC"
    utc_iso, ist_str, dt = normalize_and_parse_date(date_str)
    assert utc_iso == "2026-09-03T08:17:04Z"
    assert ist_str == "2026-09-03 13:47:04"
    assert dt is not None


def test_debt_pairs_net_to_zero_and_unbalanced_exception(sample_payment_file: Path, tmp_project: AppConfig, tmp_path: Path):
    """Verify debt pairs are validated and unbalanced debts log DEBT_PAIR_UNBALANCED."""
    res = load_payment_report(sample_payment_file, tmp_project, run_id="TEST_DEBT")
    debt_exceptions = [e for e in res.exceptions if e.reason_code == "DEBT_PAIR_UNBALANCED"]
    assert len(debt_exceptions) == 0

    # Create file with unbalanced debt
    unbalanced_rows = (
        SAMPLE_PREAMBLE_AND_HEADER +
        '"25 Aug 2026 8:17:01 am UTC","27652530972","Debt","","","Cross Account","","amazon.in","Standard","Merchant","","","","0.00","0.00","0.00","0.00","0.00","0.00","0.00","0.00","0.00","0.00","0.00","0.00","-50.00","-50.00","Released","28 Aug 2026"\n'
        '"25 Aug 2026 8:17:01 am UTC","27652530972","Debt","","","Cross Account","","amazon.in","Standard","Merchant","","","","0.00","0.00","0.00","0.00","0.00","0.00","0.00","0.00","0.00","0.00","0.00","0.00","40.00","40.00","Released","28 Aug 2026"\n'
    )
    unbalanced_f = tmp_path / "unbalanced.csv"
    unbalanced_f.write_text(unbalanced_rows, encoding="utf-8-sig")

    res_unbal = load_payment_report(unbalanced_f, tmp_project, run_id="TEST_UNBAL")
    unbal_exc = [e for e in res_unbal.exceptions if e.reason_code == "DEBT_PAIR_UNBALANCED"]
    assert len(unbal_exc) == 1
    assert unbal_exc[0].severity == "WARNING"


def test_spot_check_orders_and_pivot_values(sample_payment_file: Path, tmp_project: AppConfig):
    """Verify exact pivot calculations for the 3 spot-check orders."""
    res = load_payment_report(sample_payment_file, tmp_project, run_id="TEST_PIVOT")

    pivot_map = res.order_pivot_df.set_index("order_id").to_dict("index")

    # 1. 407-7389902-9127540: Order +371.14, Shipping -64.90 -> total_net 306.24, net_ex_shipping 371.14
    ord1 = pivot_map["407-7389902-9127540"]
    assert ord1["order_amount"] == Decimal("371.14")
    assert ord1["shipping_fee"] == Decimal("-64.90")
    assert ord1["total_net"] == Decimal("306.24")
    assert ord1["net_ex_shipping"] == Decimal("371.14")
    assert ord1["line_count"] == 2

    # 2. 171-0478052-6676325: Order +371.14, Shipping -64.90, Refund -446.66 -> total_net -140.42, net_ex_shipping -75.52
    ord2 = pivot_map["171-0478052-6676325"]
    assert ord2["order_amount"] == Decimal("371.14")
    assert ord2["shipping_fee"] == Decimal("-64.90")
    assert ord2["refund_amount"] == Decimal("-446.66")
    assert ord2["total_net"] == Decimal("-140.42")
    assert ord2["net_ex_shipping"] == Decimal("-75.52")
    assert ord2["line_count"] == 3

    # 3. 404-6736655-4486719: Order 0.00, Shipping -64.90 (in different settlement) -> total_net -64.90, net_ex_shipping 0.00
    ord3 = pivot_map["404-6736655-4486719"]
    assert ord3["order_amount"] == Decimal("0.00")
    assert ord3["shipping_fee"] == Decimal("-64.90")
    assert ord3["total_net"] == Decimal("-64.90")
    assert ord3["net_ex_shipping"] == Decimal("0.00")
    assert "27652530972" in ord3["settlement_ids"]
    assert "27788990011" in ord3["settlement_ids"]

    # Non-order rows grouped under '(no order id)'
    assert len(res.non_order_df) == 1
    assert res.non_order_df.iloc[0]["order_id"] == "(no order id)"
    # Transfer -101.44 + Debt 0.00 = -101.44
    assert res.non_order_df.iloc[0]["total_net"] == Decimal("-101.44")


def test_full_41_row_real_sample_expectations(tmp_project: AppConfig):
    """Test full 41-row sample expectations: 13 order ids, 1 '(no order id)' row, Grand Total -755.26."""
    if not REAL_PAYMENT_FILE.exists():
        pytest.skip(f"Real payment file not found at {REAL_PAYMENT_FILE}")

    res = load_payment_report(REAL_PAYMENT_FILE, tmp_project, run_id="TEST_REAL_41")

    # 41 total lines
    assert len(res.lines_df) == 41

    # 13 distinct order ids in pivot
    assert len(res.order_pivot_df) == 13
    assert res.order_pivot_df["order_id"].nunique() == 13

    # 1 '(no order id)' row for non-order rows
    assert len(res.non_order_df) == 1
    assert res.non_order_df.iloc[0]["order_id"] == "(no order id)"

    # Grand Total must equal -755.26 and Transfers -1065.72
    assert res.grand_total == Decimal("-755.26")
    transfers = res.lines_df[res.lines_df["type"].str.lower() == "transfer"]["total"].sum()
    assert transfers == Decimal("-1065.72")

    # Check the 3 spot-check orders on the real data
    pivot_map = res.order_pivot_df.set_index("order_id").to_dict("index")

    # 171-0478052-6676325: Order +371.14, Shipping -64.90, Refund -446.66 -> total_net -140.42, net_ex_shipping -75.52
    assert pivot_map["171-0478052-6676325"]["total_net"] == Decimal("-140.42")
    assert pivot_map["171-0478052-6676325"]["net_ex_shipping"] == Decimal("-75.52")

    # 404-6736655-4486719: Order 0.00, Shipping -64.90 -> total_net -64.90, net_ex_shipping 0.00
    assert pivot_map["404-6736655-4486719"]["total_net"] == Decimal("-64.90")
    assert pivot_map["404-6736655-4486719"]["net_ex_shipping"] == Decimal("0.00")

    # 407-7389902-9127540: Order +371.14, Shipping -64.90 -> total_net 306.24, net_ex_shipping 371.14
    assert pivot_map["407-7389902-9127540"]["total_net"] == Decimal("306.24")
    assert pivot_map["407-7389902-9127540"]["net_ex_shipping"] == Decimal("371.14")


def test_row_total_mismatch_exception(tmp_path: Path, tmp_project: AppConfig):
    """Verify row where component sum differs from total logs ROW_TOTAL_MISMATCH exception."""
    mismatch_row = (
        SAMPLE_PREAMBLE_AND_HEADER +
        # product_sales=380.00, total says 400.00 (mismatch)
        '"10 Aug 2026 8:17:01 am UTC","27652530972","Order","407-MISMATCH","SKU-1","T-Shirt","1","amazon.in","Standard","Merchant","DHANBAD","JHARKHAND","826004","380.00","0.00","0.00","0.00","0.00","0.00","0.00","0.00","0.00","0.00","0.00","0.00","0.00","400.00","Released","15 Aug 2026"\n'
    )
    mismatch_f = tmp_path / "mismatch.csv"
    mismatch_f.write_text(mismatch_row, encoding="utf-8-sig")

    res = load_payment_report(mismatch_f, tmp_project, run_id="TEST_MISMATCH")
    mismatch_exc = [e for e in res.exceptions if e.reason_code == "ROW_TOTAL_MISMATCH"]
    assert len(mismatch_exc) == 1
    assert mismatch_exc[0].severity == "WARNING"


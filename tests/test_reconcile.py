"""Unit tests for Phase 4: LEFT JOIN and Delivery Status Categorization (Main Output)."""

from decimal import Decimal
from pathlib import Path
import pandas as pd
import pytest

from src.config import AppConfig
from src.orders import load_order_report
from src.payments import load_payment_report
from src.reconcile import reconcile_orders_and_payments

REAL_ORDER_FILE = Path("D:/2026/NOVEMBER MONTH/13-11-2026/157818151187020770.txt")
REAL_PAYMENT_FILE = Path("D:/2026/NOVEMBER MONTH/13-11-2026/2026Aug1-2026Oct31CustomUnifiedTransaction.csv")


def test_reconcile_left_join_five_row_sample_expectations(tmp_project: AppConfig):
    """Test LEFT JOIN on 5-row order sample + 41-row payment sample matching all Section 8 specs:

    - Output contains ONLY orders present in the order report (exactly 5 rows).
    - 9 payment-only orders are excluded from the output.
    - Grand total quantity = 4.
    - Grand total of 'total' across the orders = -39.50.
    """
    if not (REAL_ORDER_FILE.exists() and REAL_PAYMENT_FILE.exists()):
        pytest.skip("Real sample files not available on test system")

    order_res = load_order_report(REAL_ORDER_FILE, tmp_project, run_id="TEST_RECON_ORD")
    payment_res = load_payment_report(REAL_PAYMENT_FILE, tmp_project, run_id="TEST_RECON_PAY")

    # 1. Test default net_ex_shipping basis
    res_ex = reconcile_orders_and_payments(
        order_res, payment_res, tmp_project, basis="net_ex_shipping", run_id="TEST_RECON_EX"
    )

    # Output row count must strictly equal 5 (number of distinct amazon-order-id in order report)
    assert len(res_ex.recon_df) == 5
    assert res_ex.excluded_payment_count == 9
    assert len(res_ex.payment_only_df) == 9

    df_ex = res_ex.recon_df.set_index("amazon-order-id")

    # Row 1: 171-0478052-6676325 | AA(ARN)BLACK_302-M | 1 | Shipped - Delivered to Buyer | -140.42 | RVP | STATUS_MISMATCH, REFUND_LOSS
    ord1 = df_ex.loc["171-0478052-6676325"]
    assert ord1["sku"] == "AA(ARN)BLACK_302-M"
    assert ord1["quantity"] == 1
    assert ord1["order_status"] == "Shipped - Delivered to Buyer"
    assert ord1["total"] == Decimal("-140.42")
    assert ord1["net_ex_shipping"] == Decimal("-75.52")
    assert ord1["delivery_status"] == "RVP"
    assert "STATUS_MISMATCH" in ord1["flags"]
    assert "REFUND_LOSS" in ord1["flags"]

    # Row 2: 404-3526450-9656333 | AA(ARN)BLACK_173-M | 1 | Shipped - Delivered to Buyer | -140.42 | RVP | STATUS_MISMATCH, REFUND_LOSS
    ord2 = df_ex.loc["404-3526450-9656333"]
    assert ord2["sku"] == "AA(ARN)BLACK_173-M"
    assert ord2["quantity"] == 1
    assert ord2["order_status"] == "Shipped - Delivered to Buyer"
    assert ord2["total"] == Decimal("-140.42")
    assert ord2["net_ex_shipping"] == Decimal("-75.52")
    assert ord2["delivery_status"] == "RVP"
    assert "STATUS_MISMATCH" in ord2["flags"]
    assert "REFUND_LOSS" in ord2["flags"]

    # Row 3: 404-6736655-4486719 | AA(ARN)BLACK_173-L | 1 | Shipped - Delivered to Buyer | -64.90 | RTO | REPLACEMENT_ORDER_COST, ZERO_PRICE_SHIPPED, STATUS_MISMATCH
    ord3 = df_ex.loc["404-6736655-4486719"]
    assert ord3["sku"] == "AA(ARN)BLACK_173-L"
    assert ord3["quantity"] == 1
    assert ord3["order_status"] == "Shipped - Delivered to Buyer"
    assert ord3["total"] == Decimal("-64.90")
    assert ord3["net_ex_shipping"] == Decimal("0.00")
    assert ord3["delivery_status"] == "RTO"
    assert "REPLACEMENT_ORDER_COST" in ord3["flags"]
    assert "ZERO_PRICE_SHIPPED" in ord3["flags"]
    assert "STATUS_MISMATCH" in ord3["flags"]

    # Row 4: 407-7389902-9127540 | AA(ARN)BLACK_43-M | 1 | Shipped - Delivered to Buyer | 306.24 | Delivered
    ord4 = df_ex.loc["407-7389902-9127540"]
    assert ord4["sku"] == "AA(ARN)BLACK_43-M"
    assert ord4["quantity"] == 1
    assert ord4["order_status"] == "Shipped - Delivered to Buyer"
    assert ord4["total"] == Decimal("306.24")
    assert ord4["net_ex_shipping"] == Decimal("371.14")
    assert ord4["delivery_status"] == "Delivered"

    # Row 5: 408-3937651-1846748 | AA(ARN)LAVANDER_380-M | 0 | Cancelled | blank | Cancelled
    ord5 = df_ex.loc["408-3937651-1846748"]
    assert ord5["sku"] == "AA(ARN)LAVANDER_380-M"
    assert ord5["quantity"] == 0
    assert ord5["order_status"] == "Cancelled"
    assert ord5["total"] is None
    assert ord5["net_ex_shipping"] is None
    assert ord5["delivery_status"] == "Cancelled"

    # Grand total quantity = 4
    assert res_ex.recon_df["quantity"].sum() == 4

    # Grand total of total = -140.42 + -140.42 + -64.90 + 306.24 = -39.50
    valid_totals = res_ex.recon_df["total"].dropna()
    assert sum(valid_totals) == Decimal("-39.50")


def test_status_counts_both_basis_settings(tmp_project: AppConfig):
    """Test status counts under both basis settings on the master order report rows:

    - net_ex_shipping basis -> RVP 2, RTO 1, Delivered 1, Cancelled 1
    - total_net basis       -> RVP 3, RTO 0, Delivered 1, Cancelled 1 (RTO row becomes RVP)
    """
    if not (REAL_ORDER_FILE.exists() and REAL_PAYMENT_FILE.exists()):
        pytest.skip("Real sample files not available on test system")

    order_res = load_order_report(REAL_ORDER_FILE, tmp_project, run_id="TEST_BASIS_ORD")
    payment_res = load_payment_report(REAL_PAYMENT_FILE, tmp_project, run_id="TEST_BASIS_PAY")

    # 1. net_ex_shipping basis
    res_ex = reconcile_orders_and_payments(
        order_res, payment_res, tmp_project, basis="net_ex_shipping", run_id="TEST_BASIS_EX"
    )
    assert res_ex.status_counts == {"RVP": 2, "Delivered": 1, "RTO": 1, "Cancelled": 1}

    # 2. total_net basis (RTO order becomes RVP)
    res_tot = reconcile_orders_and_payments(
        order_res, payment_res, tmp_project, basis="total_net", run_id="TEST_BASIS_TOT"
    )
    assert res_tot.status_counts == {"RVP": 3, "Delivered": 1, "Cancelled": 1}
    assert res_tot.status_counts.get("RTO", 0) == 0


def test_shipped_not_settled_order_only(tmp_path: Path, tmp_project: AppConfig):
    """Verify order with shipped status but no payment lines gets 'Not settled' and SHIPPED_NOT_SETTLED flag."""
    order_file = tmp_path / "order_unsettled.txt"
    from tests.test_orders import SAMPLE_HEADER, _make_tsv_row
    order_file.write_text(SAMPLE_HEADER + _make_tsv_row("ORD-UNSETTLED-01", "ITEM-01", "SKU-01", "Shipped", "1"), encoding="utf-8")

    from src.payments import PaymentParseResult
    empty_payment = PaymentParseResult(
        file_name="empty.csv",
        file_hash="h1",
        lines_df=pd.DataFrame(),
        order_pivot_df=pd.DataFrame(),
        non_order_df=pd.DataFrame(),
        grand_total=Decimal("0.00"),
    )

    order_res = load_order_report(order_file, tmp_project, run_id="TEST_UNSET")
    res = reconcile_orders_and_payments(order_res, empty_payment, tmp_project, run_id="TEST_UNSET_RECON")

    row = res.recon_df.iloc[0]
    assert row["amazon-order-id"] == "ORD-UNSETTLED-01"
    assert row["delivery_status"] == "Not settled"
    assert "SHIPPED_NOT_SETTLED" in row["flags"]

    shipped_not_settled_exc = [e for e in res.exceptions if e.reason_code == "SHIPPED_NOT_SETTLED"]
    assert len(shipped_not_settled_exc) == 1


def test_grand_total_kept_out_of_dataframe(tmp_project: AppConfig):
    """Verify logic dataframe does not inject Grand Total row; grand total is on result object."""
    if not (REAL_ORDER_FILE.exists() and REAL_PAYMENT_FILE.exists()):
        pytest.skip("Real sample files not available on test system")

    order_res = load_order_report(REAL_ORDER_FILE, tmp_project)
    payment_res = load_payment_report(REAL_PAYMENT_FILE, tmp_project)
    res = reconcile_orders_and_payments(order_res, payment_res, tmp_project)

    # Check no 'Grand Total' or 'Total' exists in amazon-order-id
    assert not res.recon_df["amazon-order-id"].str.lower().str.contains("total").any()
    # Check grand total is preserved as Decimal
    assert res.grand_total == Decimal("-755.26")

"""Unit tests for Phase 2: Load Order Report and Order-Level Summary."""

from decimal import Decimal
from pathlib import Path
import pandas as pd
import pytest

from src.config import AppConfig
from src.orders import is_order_report, load_order_report


SAMPLE_HEADER = (
    "amazon-order-id\tmerchant-order-id\tpurchase-date\tlast-updated-date\torder-status\t"
    "fulfillment-channel\tsales-channel\torder-channel\turl\tship-service-level\tproduct-name\t"
    "sku\tasin\titem-status\tquantity\tcurrency\titem-price\titem-tax\tshipping-price\t"
    "shipping-tax\tgift-wrap-price\tgift-wrap-tax\titem-promotion-discount\tship-promotion-discount\t"
    "ship-city\tship-state\tship-postal-code\tship-country\tpromotion-ids\tfulfilled-by\t"
    "amazon-programs\torder-item-id\tis-prime \n"
)


def _make_tsv_row(
    order_id: str,
    item_id: str,
    sku: str,
    status: str,
    qty: str,
    price: str = "399.0",
    tax: str = "19.0",
    currency: str = "INR",
    purchase_date: str = "2026-08-15T05:26:38+00:00",
    updated_date: str = "2026-08-25T05:42:09+00:00",
    city: str = "MUMBAI",
    state: str = "MAHARASHTRA",
    postal: str = "400001",
) -> str:
    return (
        f"{order_id}\t\t{purchase_date}\t{updated_date}\t{status}\t"
        f"Merchant\tAmazon.in\tWebsiteOrderChannel\t\tStandard\tT-Shirt\t"
        f"{sku}\tB001\tShipped\t{qty}\t{currency}\t{price}\t{tax}\t0.0\t"
        f"0.0\t0.0\t0.0\t0.0\t0.0\t"
        f"{city}\t{state}\t{postal}\tIN\t\tEasy Ship\t\t{item_id}\ttrue\n"
    )


@pytest.fixture
def five_row_order_file(tmp_path: Path) -> Path:
    """Create the exact 5-row canonical order report fixture."""
    lines = [
        SAMPLE_HEADER,
        # 1. 171-0478052-6676325 | AA(ARN)BLACK_302-M | Shipped - Delivered to Buyer | qty 1
        _make_tsv_row("171-0478052-6676325", "66268119131002", "AA(ARN)BLACK_302-M",
                      "Shipped - Delivered to Buyer", "1", "399.0", "19.0"),
        # 2. 404-3526450-9656333 | AA(ARN)BLACK_173-M | Shipped - Delivered to Buyer | qty 1
        _make_tsv_row("404-3526450-9656333", "66526370752162", "AA(ARN)BLACK_173-M",
                      "Shipped - Delivered to Buyer", "1", "399.0", "19.0"),
        # 3. 404-6736655-4486719 | AA(ARN)BLACK_173-L | Shipped - Delivered to Buyer | qty 1 | price 0.0 (ZERO_PRICE_SHIPPED)
        _make_tsv_row("404-6736655-4486719", "66808478810802", "AA(ARN)BLACK_173-L",
                      "Shipped - Delivered to Buyer", "1", "0.0", ""),
        # 4. 407-7389902-9127540 | AA(ARN)BLACK_43-M | Shipped - Delivered to Buyer | qty 1
        _make_tsv_row("407-7389902-9127540", "66535898600442", "AA(ARN)BLACK_43-M",
                      "Shipped - Delivered to Buyer", "1", "399.0", "19.0"),
        # 5. 408-3937651-1846748 | AA(ARN)LAVANDER_380-M | Cancelled | qty 0
        _make_tsv_row("408-3937651-1846748", "65983242908642", "AA(ARN)LAVANDER_380-M",
                      "Cancelled", "0", "", ""),
    ]
    file_path = tmp_path / "orders_5rows.txt"
    file_path.write_text("".join(lines), encoding="utf-8")
    return file_path


def test_is_order_report_detection(five_row_order_file: Path, tmp_path: Path):
    """Verify is_order_report detects order report headers correctly."""
    assert is_order_report(five_row_order_file) is True

    non_order_file = tmp_path / "not_order.csv"
    non_order_file.write_text("date/time,settlement id,type,order id,sku,total\n", encoding="utf-8")
    assert is_order_report(non_order_file) is False


def test_load_order_report_five_row_fixture(five_row_order_file: Path, tmp_project: AppConfig):
    """Test parsing the canonical 5-row fixture and order-level summary."""
    res = load_order_report(five_row_order_file, tmp_project, run_id="TEST_RUN_5")

    # Item-level dataframe
    assert len(res.items_df) == 5
    # Enforce strictly 1 row per order id in summary
    assert len(res.summary_df) == 5
    assert res.summary_df["amazon-order-id"].nunique() == 5

    # Check total quantity: 1 + 1 + 1 + 1 + 0 = 4
    assert res.summary_df["total_qty"].sum() == 4

    summary_map = res.summary_df.set_index("amazon-order-id").to_dict("index")

    # Order 1: 171-0478052-6676325
    ord1 = summary_map["171-0478052-6676325"]
    assert ord1["sku"] == "AA(ARN)BLACK_302-M"
    assert ord1["order_status"] == "Shipped - Delivered to Buyer"
    assert ord1["total_qty"] == 1
    assert ord1["line_count"] == 1

    # Order 2: 404-3526450-9656333
    ord2 = summary_map["404-3526450-9656333"]
    assert ord2["sku"] == "AA(ARN)BLACK_173-M"
    assert ord2["order_status"] == "Shipped - Delivered to Buyer"
    assert ord2["total_qty"] == 1

    # Order 3: 404-6736655-4486719 (price 0.0)
    ord3 = summary_map["404-6736655-4486719"]
    assert ord3["sku"] == "AA(ARN)BLACK_173-L"
    assert ord3["order_status"] == "Shipped - Delivered to Buyer"
    assert ord3["total_qty"] == 1

    # Order 4: 407-7389902-9127540
    ord4 = summary_map["407-7389902-9127540"]
    assert ord4["sku"] == "AA(ARN)BLACK_43-M"
    assert ord4["order_status"] == "Shipped - Delivered to Buyer"
    assert ord4["total_qty"] == 1

    # Order 5: 408-3937651-1846748 (Cancelled)
    ord5 = summary_map["408-3937651-1846748"]
    assert ord5["sku"] == "AA(ARN)LAVANDER_380-M"
    assert ord5["order_status"] == "Cancelled"
    assert ord5["total_qty"] == 0

    # Exceptions check: ZERO_PRICE_SHIPPED must be present for order 404-6736655-4486719
    zero_price_exc = [
        e for e in res.exceptions
        if e.reason_code == "ZERO_PRICE_SHIPPED" and e.reference_id == "66808478810802"
    ]
    assert len(zero_price_exc) == 1
    assert zero_price_exc[0].severity == "WARNING"


def test_mixed_status_detection(tmp_path: Path, tmp_project: AppConfig):
    """Test multi-item order with mixed statuses sets order_status='Mixed' and logs MIXED_STATUS."""
    lines = [
        SAMPLE_HEADER,
        _make_tsv_row("ORD-MIXED-001", "ITEM-01", "SKU-A", "Shipped - Delivered to Buyer", "1"),
        _make_tsv_row("ORD-MIXED-001", "ITEM-02", "SKU-B", "Cancelled", "0"),
    ]
    f = tmp_path / "mixed_orders.txt"
    f.write_text("".join(lines), encoding="utf-8")

    res = load_order_report(f, tmp_project, run_id="TEST_MIXED")
    assert len(res.summary_df) == 1
    row = res.summary_df.iloc[0]

    assert row["amazon-order-id"] == "ORD-MIXED-001"
    assert row["order_status"] == "Mixed"
    assert row["sku"] == "SKU-A | SKU-B"
    assert row["line_count"] == 2
    assert row["total_qty"] == 1

    mixed_exc = [e for e in res.exceptions if e.reason_code == "MIXED_STATUS"]
    assert len(mixed_exc) == 1
    assert mixed_exc[0].reference_id == "ORD-MIXED-001"


def test_duplicate_order_item_resolution(tmp_path: Path, tmp_project: AppConfig):
    """Test duplicate order-item-id keeps the latest last-updated-date record."""
    lines = [
        SAMPLE_HEADER,
        # Older version of item
        _make_tsv_row("ORD-DUP-01", "ITEM-DUP", "SKU-OLD", "Shipped", "1",
                      updated_date="2026-08-10T10:00:00+00:00"),
        # Newer version of same item
        _make_tsv_row("ORD-DUP-01", "ITEM-DUP", "SKU-NEW", "Shipped - Delivered to Buyer", "1",
                      updated_date="2026-08-20T10:00:00+00:00"),
    ]
    f = tmp_path / "dup_orders.txt"
    f.write_text("".join(lines), encoding="utf-8")

    res = load_order_report(f, tmp_project, run_id="TEST_DUP")
    # Only 1 unique item row kept
    assert len(res.items_df) == 1
    assert res.items_df.iloc[0]["sku"] == "SKU-NEW"
    assert res.items_df.iloc[0]["order_status"] == "Shipped - Delivered to Buyer"

    # Exception logged
    dup_exc = [e for e in res.exceptions if e.reason_code == "DUPLICATE_ORDER_ITEM"]
    assert len(dup_exc) == 1


def test_zero_drop_surrogates_and_validation(tmp_path: Path, tmp_project: AppConfig):
    """Test zero-drop policy: missing IDs/SKUs or bad values are assigned surrogates and kept."""
    lines = [
        SAMPLE_HEADER,
        # Missing order-item-id
        _make_tsv_row("ORD-BAD-01", "", "SKU-VALID", "Shipped", "1"),
        # Missing SKU
        _make_tsv_row("ORD-BAD-02", "ITEM-02", "", "Shipped", "1"),
        # Unexpected Currency
        _make_tsv_row("ORD-BAD-03", "ITEM-03", "SKU-USD", "Shipped", "1", currency="USD"),
    ]
    f = tmp_path / "bad_orders.txt"
    f.write_text("".join(lines), encoding="utf-8")

    res = load_order_report(f, tmp_project, run_id="TEST_BAD")
    # Zero drop: all 3 rows must be in items_df
    assert len(res.items_df) == 3
    assert len(res.summary_df) == 3

    # Check surrogate key
    row1 = res.items_df.iloc[0]
    assert row1["order_item_id"].startswith("MISSING_ORD-BAD-01")

    # Exceptions verified
    codes = {e.reason_code for e in res.exceptions}
    assert "MISSING_ORDER_ITEM_ID" in codes
    assert "MISSING_SKU" in codes
    assert "UNEXPECTED_CURRENCY" in codes


def test_raw_file_copied_untouched(tmp_path: Path, tmp_project: AppConfig):
    """Test uploaded raw file is copied untouched to /data/raw/orders."""
    upload_dir = tmp_path / "uploads"
    upload_dir.mkdir()
    f = upload_dir / "my_order_file.txt"
    f.write_text(SAMPLE_HEADER + _make_tsv_row("ORD-1", "IT-1", "SKU-1", "Shipped", "1"), encoding="utf-8")

    res = load_order_report(f, tmp_project, run_id="TEST_COPY", archive_raw=True)
    copied_file = tmp_project.orders_raw_dir / "my_order_file.txt"
    assert copied_file.exists()
    assert copied_file.read_bytes() == f.read_bytes()


"""Order report parser and order-level summary builder for Amazon.in.

Implements Phase 2:
- Tab-separated UTF-8 order report ingestion with header whitespace stripping.
- Key = order-item-id, converts UTC dates to IST.
- Blanks treated as 0 for quantity and money.
- Money stored as Decimal rounded to 2 decimal places.
- Privacy: keeps only ship-state and ship-postal-code.
- Order-level summary: amazon-order-id, total_qty, sku, order_status, line_count.
- Mixed status detection -> "Mixed" + MIXED_STATUS exception.
- Enforces strictly 1 row per amazon-order-id.
- Captures exceptions without dropping rows:
  ZERO_PRICE_SHIPPED, MISSING_SKU, DUPLICATE_ORDER_ITEM, UNEXPECTED_CURRENCY,
  UNKNOWN_STATUS, MISSING_ORDER_ITEM_ID.
"""

from dataclasses import dataclass, field
from datetime import datetime
from decimal import Decimal, InvalidOperation
import hashlib
import logging
from pathlib import Path
import shutil
from typing import Any, Dict, List, Optional, Tuple
from zoneinfo import ZoneInfo
from dateutil import parser as date_parser
import pandas as pd

from src.config import AppConfig

logger = logging.getLogger("amazon_recon.orders")


@dataclass
class OrderException:
    """Exception record for an order or order-item row."""

    run_id: str
    source: str
    reference_id: Optional[str]
    reason_code: str
    severity: str
    details: str


@dataclass
class OrderParseResult:
    """Result of parsing an order report file."""

    file_name: str
    file_hash: str
    items_df: pd.DataFrame
    summary_df: pd.DataFrame
    exceptions: List[OrderException] = field(default_factory=list)


def compute_file_hash(file_path: Path) -> str:
    """Compute SHA-256 hash of a file for idempotency checking."""
    hasher = hashlib.sha256()
    with open(file_path, "rb") as f:
        while chunk := f.read(65536):
            hasher.update(chunk)
    return hasher.hexdigest()


def is_order_report(file_path: Path) -> bool:
    """Check if a file is an Amazon Order report by inspecting headers.

    Order report contains both 'amazon-order-id' and 'order-item-id'.
    """
    try:
        with open(file_path, "r", encoding="utf-8", errors="replace") as f:
            first_line = f.readline()
        tokens = [t.strip().lower() for t in first_line.split("\t")]
        return "amazon-order-id" in tokens and "order-item-id" in tokens
    except Exception:
        return False


def parse_decimal(val: Any) -> Decimal:
    """Parse a value into a 2-decimal rounded Decimal (defaulting to 0.00)."""
    if val is None or pd.isna(val):
        return Decimal("0.00")
    s = str(val).strip().replace(",", "")
    if not s or s.lower() == "nan":
        return Decimal("0.00")
    try:
        return Decimal(s).quantize(Decimal("0.01"))
    except InvalidOperation:
        return Decimal("0.00")


def parse_int(val: Any) -> int:
    """Parse an integer value, defaulting to 0 for blanks/NaN."""
    if val is None or pd.isna(val):
        return 0
    s = str(val).strip()
    if not s or s.lower() == "nan":
        return 0
    try:
        return int(float(s))
    except (ValueError, TypeError):
        return 0


def clean_str(val: Any) -> Optional[str]:
    """Clean string value, returning None if empty or NaN."""
    if val is None or pd.isna(val):
        return None
    s = str(val).strip()
    return s if s and s.lower() != "nan" else None


def parse_datetime_to_ist(date_str: Optional[str], tz_name: str = "Asia/Kolkata") -> Tuple[Optional[str], Optional[str], Optional[datetime]]:
    """Parse UTC date string and return UTC ISO, IST formatted string, and IST datetime object."""
    if not date_str:
        return None, None, None
    s = str(date_str).strip()
    if not s or s.lower() == "nan":
        return None, None, None

    try:
        dt = date_parser.parse(s)
        # If naive, assume UTC
        if dt.tzinfo is None:
            dt = dt.replace(tzinfo=ZoneInfo("UTC"))
        ist_tz = ZoneInfo(tz_name)
        dt_ist = dt.astimezone(ist_tz)
        utc_iso = dt.strftime("%Y-%m-%dT%H:%M:%SZ")
        ist_str = dt_ist.strftime("%Y-%m-%d %H:%M:%S")
        return utc_iso, ist_str, dt_ist
    except Exception:
        return s, s, None


def load_order_report(
    file_path: Path | str,
    config: AppConfig,
    run_id: str = "DEFAULT_RUN",
    archive_raw: bool = False,
) -> OrderParseResult:
    """Parse an Amazon order report into an item-level dataframe and an order-level summary.

    Args:
        file_path: Path to order report (.txt / .tsv).
        config: Application configuration.
        run_id: Identifier for current execution.
        archive_raw: If True, copies file to config.orders_raw_dir.

    Returns:
        OrderParseResult containing items_df, summary_df, and exceptions.
    """
    path = Path(file_path).resolve()
    if not path.exists():
        raise FileNotFoundError(f"Order report file not found: {path}")

    if archive_raw:
        config.ensure_directories()
        dest_path = config.orders_raw_dir / path.name
        if dest_path.resolve() != path:
            shutil.copy2(path, dest_path)
            logger.info(f"Copied raw order report to {dest_path}")
        target_path = dest_path
    else:
        target_path = path

    file_hash = compute_file_hash(target_path)
    file_name = target_path.name
    exceptions: List[OrderException] = []

    # Read tab-separated file with header whitespace stripping
    with open(target_path, "r", encoding="utf-8", errors="replace") as f:
        header_line = f.readline()
    raw_headers = [h.strip() for h in header_line.split("\t")]

    df = pd.read_csv(
        target_path,
        sep="\t",
        names=raw_headers,
        header=0,
        dtype=str,
        keep_default_na=False,
        encoding="utf-8",
        on_bad_lines="warn",
    )

    items_records: List[Dict[str, Any]] = []
    seen_order_items: Dict[str, Tuple[int, datetime]] = {}  # order_item_id -> (index in items_records, updated_dt)

    for idx, row in df.iterrows():
        source_row = int(idx) + 2

        amazon_order_id = clean_str(row.get("amazon-order-id"))
        raw_order_item_id = clean_str(row.get("order-item-id"))
        sku = clean_str(row.get("sku"))
        order_status = clean_str(row.get("order-status")) or "UNKNOWN"
        currency = clean_str(row.get("currency")) or "INR"

        # Missing order-item-id handling (surrogate key to prevent row drop)
        if not raw_order_item_id:
            order_item_id = f"MISSING_{amazon_order_id or 'NO_ORD'}_{source_row}"
            exceptions.append(OrderException(
                run_id=run_id,
                source="order",
                reference_id=amazon_order_id,
                reason_code="MISSING_ORDER_ITEM_ID",
                severity="CRITICAL",
                details=f"Missing order-item-id at row {source_row}. Assigned surrogate: {order_item_id}",
            ))
        else:
            order_item_id = raw_order_item_id

        # Missing SKU check
        if not sku:
            exceptions.append(OrderException(
                run_id=run_id,
                source="order",
                reference_id=order_item_id,
                reason_code="MISSING_SKU",
                severity="ERROR",
                details=f"Missing SKU for order {amazon_order_id} at row {source_row}",
            ))

        # Unexpected currency check
        if currency != "INR":
            exceptions.append(OrderException(
                run_id=run_id,
                source="order",
                reference_id=order_item_id,
                reason_code="UNEXPECTED_CURRENCY",
                severity="WARNING",
                details=f"Unexpected currency '{currency}' for order {amazon_order_id} (expected INR)",
            ))

        # Unknown order status check
        if order_status not in config.order_status_map:
            exceptions.append(OrderException(
                run_id=run_id,
                source="order",
                reference_id=order_item_id,
                reason_code="UNKNOWN_STATUS",
                severity="WARNING",
                details=f"Unrecognized order status '{order_status}' for order {amazon_order_id}",
            ))

        # Dates
        purchase_utc, purchase_ist, _ = parse_datetime_to_ist(row.get("purchase-date"), config.timezone)
        updated_utc, updated_ist, updated_dt = parse_datetime_to_ist(row.get("last-updated-date"), config.timezone)
        effective_updated_dt = updated_dt or datetime.min

        # Money and Quantity
        raw_price = parse_decimal(row.get("item-price"))
        item_tax = parse_decimal(row.get("item-tax"))
        quantity = parse_int(row.get("quantity"))

        if config.item_price_includes_tax:
            item_price_gross = raw_price
            item_price_net = (item_price_gross - item_tax) if item_price_gross >= item_tax else Decimal("0.00")
        else:
            item_price_net = raw_price
            item_price_gross = item_price_net + item_tax

        # Zero price check for shipped orders
        if quantity > 0 and item_price_gross == Decimal("0.00") and order_status.lower().startswith("shipped"):
            exceptions.append(OrderException(
                run_id=run_id,
                source="order",
                reference_id=order_item_id,
                reason_code="ZERO_PRICE_SHIPPED",
                severity="WARNING",
                details=f"Shipped item with price 0.00 for order {amazon_order_id} (Replacement/Warranty)",
            ))

        item_rec = {
            "order_item_id": order_item_id,
            "amazon_order_id": amazon_order_id or "UNKNOWN",
            "merchant_order_id": clean_str(row.get("merchant-order-id")),
            "purchase_date_utc": purchase_utc,
            "purchase_date_ist": purchase_ist,
            "last_updated_date_utc": updated_utc,
            "last_updated_date_ist": updated_ist,
            "order_status": order_status,
            "fulfillment_channel": clean_str(row.get("fulfillment-channel")),
            "sales_channel": clean_str(row.get("sales-channel")),
            "order_channel": clean_str(row.get("order-channel")),
            "ship_service_level": clean_str(row.get("ship-service-level")),
            "product_name": clean_str(row.get("product-name")),
            "sku": sku or "",
            "asin": clean_str(row.get("asin")),
            "item_status": clean_str(row.get("item-status")),
            "quantity": quantity,
            "currency": currency,
            "item_price_gross": item_price_gross,
            "item_price_net": item_price_net,
            "item_tax": item_tax,
            "shipping_price": parse_decimal(row.get("shipping-price")),
            "shipping_tax": parse_decimal(row.get("shipping-tax")),
            "gift_wrap_price": parse_decimal(row.get("gift-wrap-price")),
            "gift_wrap_tax": parse_decimal(row.get("gift-wrap-tax")),
            "item_promotion_discount": parse_decimal(row.get("item-promotion-discount")),
            "ship_promotion_discount": parse_decimal(row.get("ship-promotion-discount")),
            "ship_city": clean_str(row.get("ship-city")),
            "ship_state": clean_str(row.get("ship-state")),
            "ship_postal_code": clean_str(row.get("ship-postal-code")),
            "ship_country": clean_str(row.get("ship-country")),
            "promotion_ids": clean_str(row.get("promotion-ids")),
            "fulfilled_by": clean_str(row.get("fulfilled-by")),
            "amazon_programs": clean_str(row.get("amazon-programs")),
            "is_prime": clean_str(row.get("is-prime")),
            "source_file": file_name,
            "source_row": source_row,
            "file_hash": file_hash,
        }

        # Deduplication check for order-item-id
        if order_item_id in seen_order_items:
            prev_idx, prev_updated_dt = seen_order_items[order_item_id]
            exceptions.append(OrderException(
                run_id=run_id,
                source="order",
                reference_id=order_item_id,
                reason_code="DUPLICATE_ORDER_ITEM",
                severity="WARNING",
                details=f"Duplicate order-item-id '{order_item_id}'. Resolved using latest update timestamp.",
            ))
            # If current row is newer, replace previous record
            if effective_updated_dt >= prev_updated_dt:
                items_records[prev_idx] = item_rec
                seen_order_items[order_item_id] = (prev_idx, effective_updated_dt)
        else:
            new_idx = len(items_records)
            items_records.append(item_rec)
            seen_order_items[order_item_id] = (new_idx, effective_updated_dt)

    items_df = pd.DataFrame(items_records)

    # -----------------------------------------------------------------------
    # Order-Level Summary Construction
    # -----------------------------------------------------------------------
    if items_df.empty:
        summary_df = pd.DataFrame(
            columns=[
                "amazon-order-id",
                "purchase_date",
                "product_name",
                "sku",
                "total_qty",
                "item_price",
                "item_tax",
                "shipping_price",
                "ship_city",
                "ship_state",
                "ship_postal_code",
                "order_status",
                "line_count",
            ]
        )
        return OrderParseResult(file_name, file_hash, items_df, summary_df, exceptions)

    # Group by amazon_order_id to build summary
    grouped = items_df.groupby("amazon_order_id", sort=False)
    summary_records: List[Dict[str, Any]] = []

    for order_id, group in grouped:
        total_qty = int(group["quantity"].sum())

        # Dates: earliest IST purchase date
        dates = [str(d) for d in group["purchase_date_ist"].dropna() if str(d).strip()]
        purchase_date = sorted(dates)[0] if dates else ""

        # Product name: first available
        prod_names = [str(p) for p in group["product_name"].dropna() if str(p).strip()]
        product_name = prod_names[0] if prod_names else ""

        # SKUs: distinct, sorted, joined with " | "
        distinct_skus = sorted({s for s in group["sku"].dropna() if s.strip()})
        joined_sku = " | ".join(distinct_skus) if distinct_skus else ""

        # Financial values from order report
        item_price_sum = sum(group["item_price_gross"]) if "item_price_gross" in group else Decimal("0.00")
        item_tax_sum = sum(group["item_tax"]) if "item_tax" in group else Decimal("0.00")
        ship_price_sum = sum(group["shipping_price"]) if "shipping_price" in group else Decimal("0.00")

        # Geo fields
        cities = [str(c) for c in group["ship_city"].dropna() if str(c).strip()]
        ship_city = cities[0] if cities else ""
        states = [str(s) for s in group["ship_state"].dropna() if str(s).strip()]
        ship_state = states[0] if states else ""
        zips = [str(z) for z in group["ship_postal_code"].dropna() if str(z).strip()]
        ship_postal_code = zips[0] if zips else ""

        # Status: check if multiple distinct statuses
        distinct_statuses = sorted({st for st in group["order_status"].dropna() if st.strip()})
        if len(distinct_statuses) > 1:
            final_status = "Mixed"
            exceptions.append(OrderException(
                run_id=run_id,
                source="order",
                reference_id=order_id,
                reason_code="MIXED_STATUS",
                severity="WARNING",
                details=f"Order {order_id} has mixed line statuses: {distinct_statuses}",
            ))
        elif len(distinct_statuses) == 1:
            final_status = distinct_statuses[0]
        else:
            final_status = "UNKNOWN"

        line_count = len(group)

        summary_records.append({
            "amazon-order-id": order_id,
            "purchase_date": purchase_date,
            "product_name": product_name,
            "sku": joined_sku,
            "total_qty": total_qty,
            "item_price": item_price_sum,
            "item_tax": item_tax_sum,
            "shipping_price": ship_price_sum,
            "ship_city": ship_city,
            "ship_state": ship_state,
            "ship_postal_code": ship_postal_code,
            "order_status": final_status,
            "line_count": line_count,
        })

    summary_df = pd.DataFrame(summary_records)

    # Enforce strictly one row per amazon-order-id
    if summary_df["amazon-order-id"].duplicated().any():
        dups = summary_df[summary_df["amazon-order-id"].duplicated()]["amazon-order-id"].tolist()
        raise ValueError(f"Enforce one row per amazon-order-id failed; duplicates found: {dups}")

    logger.info(
        f"Parsed order report '{file_name}': {len(items_df)} item rows, "
        f"{len(summary_df)} orders, {len(exceptions)} exceptions."
    )

    return OrderParseResult(
        file_name=file_name,
        file_hash=file_hash,
        items_df=items_df,
        summary_df=summary_df,
        exceptions=exceptions,
    )


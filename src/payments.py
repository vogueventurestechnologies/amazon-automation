"""Payment report parser and order-level payment pivot builder for Amazon.in.

Implements Phase 3:
- Ingestion of Amazon 'Custom Unified Transaction' CSV (utf-8-sig).
- Dynamic header detection starting with 'date/time' (no hardcoded skiprows).
- Normalizes 'Sept' -> 'Sep' in date strings.
- Copies raw file untouched to /data/raw/payments/.
- Validates row totals against sum of component columns (else ROW_TOTAL_MISMATCH).
- Verifies Debt pairs net to 0 (else DEBT_PAIR_UNBALANCED).
- Groups non-order rows (Transfer, Debt, Service Fee) under '(no order id)'.
- Builds order-level pivot:
    order_id, total_net, line_count, order_amount, shipping_fee, refund_amount, net_ex_shipping.
- Computes overall Grand Total = sum of every row's total.
"""

from dataclasses import dataclass, field
from datetime import datetime
from decimal import Decimal, InvalidOperation
import hashlib
import logging
from pathlib import Path
import shutil
from typing import Any, Dict, List, Optional, Set, Tuple
from zoneinfo import ZoneInfo
from dateutil import parser as date_parser
import pandas as pd

from src.config import AppConfig

logger = logging.getLogger("amazon_recon.payments")


@dataclass
class PaymentException:
    """Exception record for a payment report transaction row."""

    run_id: str
    source: str
    reference_id: Optional[str]
    reason_code: str
    severity: str
    details: str


@dataclass
class PaymentParseResult:
    """Result of parsing an Amazon payment report file."""

    file_name: str
    file_hash: str
    lines_df: pd.DataFrame
    order_pivot_df: pd.DataFrame
    non_order_df: pd.DataFrame
    grand_total: Decimal
    exceptions: List[PaymentException] = field(default_factory=list)


def compute_file_hash(file_path: Path) -> str:
    """Compute SHA-256 hash of a file for idempotency checking."""
    hasher = hashlib.sha256()
    with open(file_path, "rb") as f:
        while chunk := f.read(65536):
            hasher.update(chunk)
    return hasher.hexdigest()


def is_payment_report(file_path: Path) -> bool:
    """Check if a file is an Amazon Custom Unified Transaction payment report.

    Header or preamble contains 'settlement id' and ('tcs-igst' or 'tcs-cgst' or 'date/time').
    """
    try:
        with open(file_path, "r", encoding="utf-8-sig", errors="replace") as f:
            for _ in range(35):
                line = f.readline()
                if not line:
                    break
                stripped = line.strip().lower()
                if "settlement id" in stripped and ("tcs" in stripped or "date/time" in stripped or "selling fees" in stripped):
                    return True
        return False
    except Exception:
        return False


def find_header_row_index(file_path: Path) -> int:
    """Dynamically scan file to find the 0-based line index of the header row.

    The header is the first line starting with 'date/time' (or '"date/time"')
    and containing 'settlement id'.
    """
    with open(file_path, "r", encoding="utf-8-sig", errors="replace") as f:
        for idx, line in enumerate(f):
            stripped = line.strip().lower()
            if (stripped.startswith("date/time") or stripped.startswith('"date/time"')) and "settlement id" in stripped:
                return idx
    raise ValueError(f"Could not find valid payment report header row starting with 'date/time' in {file_path}")


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


def normalize_and_parse_date(
    date_str: Optional[str], tz_name: str = "Asia/Kolkata"
) -> Tuple[Optional[str], Optional[str], Optional[datetime]]:
    """Normalize Amazon date strings (e.g. '3 Sept 2026') and parse to IST/UTC."""
    if not date_str:
        return None, None, None
    s = str(date_str).strip()
    if not s or s.lower() == "nan":
        return None, None, None

    # Normalize 'Sept' -> 'Sep' before parsing
    normalized = s.replace("Sept", "Sep").replace("sept", "sep")
    try:
        dt = date_parser.parse(normalized)
        if dt.tzinfo is None:
            dt = dt.replace(tzinfo=ZoneInfo("UTC"))
        ist_tz = ZoneInfo(tz_name)
        dt_ist = dt.astimezone(ist_tz)
        utc_iso = dt.strftime("%Y-%m-%dT%H:%M:%SZ")
        ist_str = dt_ist.strftime("%Y-%m-%d %H:%M:%S")
        return utc_iso, ist_str, dt_ist
    except Exception:
        return s, s, None


def load_payment_report(
    file_path: Path | str,
    config: AppConfig,
    run_id: str = "DEFAULT_RUN",
    archive_raw: bool = False,
) -> PaymentParseResult:
    """Parse an Amazon Custom Unified Transaction payment report into an order pivot and non-order groups.

    Args:
        file_path: Path to payment report (.csv).
        config: Application configuration.
        run_id: Identifier for current execution.
        archive_raw: If True, copies file to config.payments_raw_dir.

    Returns:
        PaymentParseResult containing lines_df, order_pivot_df, non_order_df, grand_total, exceptions.
    """
    path = Path(file_path).resolve()
    if not path.exists():
        raise FileNotFoundError(f"Payment report file not found: {path}")

    if archive_raw:
        config.ensure_directories()
        dest_path = config.payments_raw_dir / path.name
        if dest_path.resolve() != path:
            shutil.copy2(path, dest_path)
            logger.info(f"Copied raw payment report to {dest_path}")
        target_path = dest_path
    else:
        target_path = path

    file_hash = compute_file_hash(target_path)
    file_name = target_path.name
    exceptions: List[PaymentException] = []

    # Find preamble header line dynamically
    header_idx = find_header_row_index(target_path)
    logger.info(f"Detected payment header row at line {header_idx} in '{file_name}'")

    raw_df = pd.read_csv(
        target_path,
        skiprows=header_idx,
        encoding="utf-8-sig",
        dtype=str,
        keep_default_na=False,
        on_bad_lines="warn",
    )

    # Strip column headers
    raw_df.columns = [c.strip() for c in raw_df.columns]

    # Map column names case-insensitively
    col_lookup = {c.lower(): c for c in raw_df.columns}

    def get_val(row: pd.Series, *candidates: str) -> Any:
        for cand in candidates:
            cand_l = cand.lower()
            if cand_l in col_lookup:
                return row.get(col_lookup[cand_l])
        return None

    # Component columns for row total validation
    component_col_keys = [
        ["product sales"],
        ["shipping credits"],
        ["gift wrap credits"],
        ["promotional rebates"],
        ["total sales tax liable(gst before adjusting tcs)", "total sales tax liable", "gst_total"],
        ["tcs-cgst"],
        ["tcs-sgst"],
        ["tcs-igst"],
        ["tds (section 194-o)", "tds"],
        ["selling fees"],
        ["fba fees"],
        ["other transaction fees"],
        ["other"],
    ]

    parsed_lines: List[Dict[str, Any]] = []
    tolerance = config.amount_tolerance

    for idx, row in raw_df.iterrows():
        source_row = header_idx + int(idx) + 2

        tx_date_raw = get_val(row, "date/time", "date")
        settlement_id = clean_str(get_val(row, "settlement id")) or "UNKNOWN_SETTLEMENT"
        tx_type = clean_str(get_val(row, "type")) or "Other"
        raw_order_id = clean_str(get_val(row, "order id"))
        sku = clean_str(get_val(row, "sku"))
        description = clean_str(get_val(row, "description"))
        quantity = parse_int(get_val(row, "quantity"))
        total = parse_decimal(get_val(row, "total"))

        utc_iso, ist_str, _ = normalize_and_parse_date(tx_date_raw, config.timezone)

        # Validate component sum against total for non-transfer rows
        component_sum = Decimal("0.00")
        has_any_component = False
        comp_dict = {}

        for key_candidates in component_col_keys:
            v_str = get_val(row, *key_candidates)
            v_dec = parse_decimal(v_str)
            if v_str and str(v_str).strip() != "":
                has_any_component = True
            component_sum += v_dec
            comp_dict[key_candidates[0]] = v_dec

        # Check row total mismatch
        if has_any_component and abs(component_sum - total) > tolerance:
            diff = total - component_sum
            exceptions.append(PaymentException(
                run_id=run_id,
                source="payment",
                reference_id=raw_order_id or settlement_id,
                reason_code="ROW_TOTAL_MISMATCH",
                severity="WARNING",
                details=(
                    f"Row {source_row} ({tx_type}) total {total} differs from sum of components "
                    f"{component_sum} by {diff}"
                ),
            ))

        line_rec = {
            "source_row": source_row,
            "settlement_id": settlement_id,
            "tx_date_utc": utc_iso,
            "tx_date_ist": ist_str,
            "type": tx_type,
            "order_id": raw_order_id,
            "clean_order_id": raw_order_id or "",
            "sku": sku or "",
            "description": description or "",
            "quantity": quantity,
            "total": total,
            "product_sales": comp_dict.get("product sales", Decimal("0.00")),
            "shipping_credits": comp_dict.get("shipping credits", Decimal("0.00")),
            "gift_wrap_credits": comp_dict.get("gift wrap credits", Decimal("0.00")),
            "promotional_rebates": comp_dict.get("promotional rebates", Decimal("0.00")),
            "gst_total": comp_dict.get("total sales tax liable(gst before adjusting tcs)", Decimal("0.00")),
            "tcs_cgst": comp_dict.get("tcs-cgst", Decimal("0.00")),
            "tcs_sgst": comp_dict.get("tcs-sgst", Decimal("0.00")),
            "tcs_igst": comp_dict.get("tcs-igst", Decimal("0.00")),
            "tds": comp_dict.get("tds (section 194-o)", Decimal("0.00")),
            "selling_fees": comp_dict.get("selling fees", Decimal("0.00")),
            "fba_fees": comp_dict.get("fba fees", Decimal("0.00")),
            "other_transaction_fees": comp_dict.get("other transaction fees", Decimal("0.00")),
            "other": comp_dict.get("other", Decimal("0.00")),
        }
        parsed_lines.append(line_rec)

    lines_df = pd.DataFrame(parsed_lines)
    grand_total = lines_df["total"].sum() if not lines_df.empty else Decimal("0.00")

    # -----------------------------------------------------------------------
    # Validate Debt Pairs
    # -----------------------------------------------------------------------
    debt_lines = lines_df[lines_df["type"].str.lower() == "debt"]
    if not debt_lines.empty:
        debt_sum = debt_lines["total"].sum()
        if abs(debt_sum) > tolerance:
            exceptions.append(PaymentException(
                run_id=run_id,
                source="payment",
                reference_id="DEBT_PAIRS",
                reason_code="DEBT_PAIR_UNBALANCED",
                severity="WARNING",
                details=f"Debt pairs do not net to 0: total net is {debt_sum}",
            ))

    # -----------------------------------------------------------------------
    # Separate Order Rows vs Non-Order Rows
    # -----------------------------------------------------------------------
    order_mask = lines_df["clean_order_id"] != ""
    order_lines_df = lines_df[order_mask].copy()
    non_order_lines_df = lines_df[~order_mask].copy()

    # Non-order rows grouped under '(no order id)'
    if not non_order_lines_df.empty:
        non_order_summary = {
            "order_id": "(no order id)",
            "total_net": non_order_lines_df["total"].sum(),
            "line_count": len(non_order_lines_df),
            "order_amount": Decimal("0.00"),
            "shipping_fee": Decimal("0.00"),
            "refund_amount": Decimal("0.00"),
            "net_ex_shipping": non_order_lines_df["total"].sum(),
            "types_breakdown": non_order_lines_df["type"].value_counts().to_dict(),
        }
        non_order_df = pd.DataFrame([non_order_summary])
    else:
        non_order_df = pd.DataFrame(
            columns=["order_id", "total_net", "line_count", "order_amount", "shipping_fee", "refund_amount", "net_ex_shipping"]
        )

    # -----------------------------------------------------------------------
    # Order-Level Payment Pivot Construction
    # -----------------------------------------------------------------------
    if order_lines_df.empty:
        order_pivot_df = pd.DataFrame(
            columns=["order_id", "total_net", "line_count", "order_amount", "shipping_fee", "refund_amount", "net_ex_shipping"]
        )
        return PaymentParseResult(
            file_name=file_name,
            file_hash=file_hash,
            lines_df=lines_df,
            order_pivot_df=order_pivot_df,
            non_order_df=non_order_df,
            grand_total=grand_total,
            exceptions=exceptions,
        )

    pivot_records: List[Dict[str, Any]] = []

    for order_id, group in order_lines_df.groupby("clean_order_id", sort=False):
        total_net = group["total"].sum()
        line_count = len(group)

        # Breakdowns by type
        order_amount = Decimal("0.00")
        shipping_fee = Decimal("0.00")
        refund_amount = Decimal("0.00")

        for _, r in group.iterrows():
            t_lower = str(r["type"]).strip().lower()
            amt = r["total"]
            if t_lower == "order":
                order_amount += amt
            elif t_lower == "shipping services":
                shipping_fee += amt
            elif t_lower == "refund":
                refund_amount += amt

        # net_ex_shipping: total excluding Shipping Services lines
        # Since shipping_fee is a negative cost, subtracting it gives total net without shipping deduction
        net_ex_shipping = total_net - shipping_fee

        # Extract representative SKU and quantity from Order lines (or Refund if no Order line)
        order_sub = group[group["type"].str.lower() == "order"]
        if not order_sub.empty:
            rep_sku = order_sub["sku"].iloc[0] or ""
            rep_qty = int(order_sub["quantity"].sum())
        else:
            rep_sku = group["sku"].iloc[0] or ""
            rep_qty = int(group["quantity"].sum())

        settlement_ids = " | ".join(sorted(set(group["settlement_id"].dropna())))

        pivot_records.append({
            "order_id": order_id,
            "total_net": total_net,
            "line_count": line_count,
            "order_amount": order_amount,
            "shipping_fee": shipping_fee,
            "refund_amount": refund_amount,
            "net_ex_shipping": net_ex_shipping,
            "sku": rep_sku,
            "quantity": rep_qty,
            "settlement_ids": settlement_ids,
        })

    order_pivot_df = pd.DataFrame(pivot_records)

    logger.info(
        f"Parsed payment report '{file_name}': {len(lines_df)} total lines, "
        f"{len(order_pivot_df)} order pivot rows, {len(non_order_lines_df)} non-order lines, "
        f"Grand Total: {grand_total}, Exceptions: {len(exceptions)}"
    )

    return PaymentParseResult(
        file_name=file_name,
        file_hash=file_hash,
        lines_df=lines_df,
        order_pivot_df=order_pivot_df,
        non_order_df=non_order_df,
        grand_total=grand_total,
        exceptions=exceptions,
    )


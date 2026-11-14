"""End-to-End Orchestration Pipeline for Amazon.in Order + Payment Reconciliation.

Handles:
1. Dynamic header-based file classification (Order vs Payment).
2. Untouched archiving to /data/raw/{orders,payments}.
3. Idempotent multi-file loading, row-hash deduplication, and date filtering.
4. Full join & delivery status reconciliation.
5. 5-Sheet Excel workbook generation.
6. Inbox file lifecycle management (archiving to /data/processed/).
7. Formatted terminal summary dashboard.
"""

from dataclasses import dataclass
from datetime import datetime
from decimal import Decimal
import hashlib
import logging
from pathlib import Path
import shutil
from typing import Any, Dict, List, Optional, Tuple
import pandas as pd

from src.config import AppConfig
from src.orders import OrderException, OrderParseResult, is_order_report, load_order_report
from src.payments import PaymentException, PaymentParseResult, is_payment_report, load_payment_report
from src.reconcile import ReconciliationResult, reconcile_orders_and_payments
from src.reporting.excel_export import generate_reconciliation_excel_report

logger = logging.getLogger("amazon_recon.pipeline")


@dataclass
class PipelineResult:
    """Summary of reconciliation pipeline execution."""
    order_files: List[Path]
    payment_files: List[Path]
    total_order_items: int
    unique_orders: int
    total_payment_lines: int
    settled_orders: int
    non_order_lines: int
    grand_total: Decimal
    closing_balance: Decimal
    delivery_status_counts: Dict[str, int]
    delivery_status_amounts: Dict[str, Decimal]
    exceptions_count: int
    exception_severities: Dict[str, int]
    report_path: Path
    basis_used: str
    from_date: Optional[str]
    to_date: Optional[str]
    excluded_payment_count: int = 0
    recon_res: Optional[ReconciliationResult] = None


def compute_file_md5(file_path: Path) -> str:
    """Compute MD5 checksum of a file."""
    hasher = hashlib.md5()
    with open(file_path, "rb") as f:
        for chunk in iter(lambda: f.read(65536), b""):
            hasher.update(chunk)
    return hasher.hexdigest()


def copy_to_raw_untouched(src_path: Path, dest_dir: Path) -> Path:
    """Copy an uploaded file untouched to raw storage, avoiding overwriting identical files."""
    dest_dir.mkdir(parents=True, exist_ok=True)
    target_path = dest_dir / src_path.name

    if target_path.exists():
        src_hash = compute_file_md5(src_path)
        target_hash = compute_file_md5(target_path)
        if src_hash == target_hash:
            logger.debug(f"File {src_path.name} already exists in raw directory with identical hash. Keeping existing.")
            return target_path
        # Name collision with different content: append timestamp
        timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
        target_path = dest_dir / f"{src_path.stem}_{timestamp}{src_path.suffix}"

    shutil.copy2(src_path, target_path)
    logger.info(f"Archived raw file untouched: {src_path} -> {target_path}")
    return target_path


def classify_files(file_paths: List[Path | str]) -> Tuple[List[Path], List[Path], List[Path]]:
    """Auto-detect order reports, payment reports, and unclassified files based on headers."""
    order_files: List[Path] = []
    payment_files: List[Path] = []
    unrecognized: List[Path] = []

    for f in file_paths:
        p = Path(f)
        if not p.exists() or not p.is_file():
            logger.warning(f"File does not exist or is not a file: {p}")
            continue

        if is_order_report(p):
            order_files.append(p)
            logger.info(f"Classified '{p.name}' as Amazon Order Report")
        elif is_payment_report(p):
            payment_files.append(p)
            logger.info(f"Classified '{p.name}' as Amazon Payment Report")
        else:
            unrecognized.append(p)
            logger.warning(f"File '{p.name}' could not be identified as Order or Payment report.")

    return order_files, payment_files, unrecognized


def load_multiple_order_reports(
    files: List[Path],
    config: AppConfig,
    from_date: Optional[str] = None,
    to_date: Optional[str] = None,
    run_id: str = "MULTI_ORD",
) -> OrderParseResult:
    """Load, combine, and deduplicate multiple order reports idempotently."""
    if not files:
        empty_items = pd.DataFrame()
        empty_summary = pd.DataFrame(columns=["amazon-order-id", "total_qty", "sku", "order_status", "line_count"])
        return OrderParseResult("NONE", "", empty_items, empty_summary, [])

    all_items: List[pd.DataFrame] = []
    all_exceptions: List[OrderException] = []
    seen_file_hashes: set[str] = set()

    for f in files:
        raw_copy = copy_to_raw_untouched(f, config.orders_raw_dir)
        parse_res = load_order_report(raw_copy, config, run_id=run_id)
        if parse_res.file_hash in seen_file_hashes:
            logger.info(f"Skipping duplicate order file with hash {parse_res.file_hash}: {f.name}")
            continue
        seen_file_hashes.add(parse_res.file_hash)
        if not parse_res.items_df.empty:
            all_items.append(parse_res.items_df)
        all_exceptions.extend(parse_res.exceptions)

    if not all_items:
        empty_summary = pd.DataFrame(columns=["amazon-order-id", "total_qty", "sku", "order_status", "line_count"])
        return OrderParseResult("COMBINED", "", pd.DataFrame(), empty_summary, all_exceptions)

    combined_items = pd.concat(all_items, ignore_index=True)

    # Deduplicate items on order-item-id keeping the latest last-updated-date
    if "order-item-id" in combined_items.columns and not combined_items.empty:
        # Sort by last-updated-date if available
        if "last-updated-date" in combined_items.columns:
            combined_items = combined_items.sort_values(by="last-updated-date", ascending=True)
        initial_len = len(combined_items)
        combined_items = combined_items.drop_duplicates(subset=["order-item-id"], keep="last")
        dedup_count = initial_len - len(combined_items)
        if dedup_count > 0:
            logger.info(f"Deduplicated {dedup_count} overlapping order item rows across input files.")

    # Apply date filter if specified (based on purchase_date_ist)
    if (from_date or to_date) and not combined_items.empty and "purchase_date_ist" in combined_items.columns:
        valid_dates = combined_items["purchase_date_ist"].dropna()
        if not valid_dates.empty:
            date_str_series = valid_dates.astype(str).str[:10]
            mask = pd.Series(True, index=combined_items.index)
            if from_date:
                mask = mask & (date_str_series >= from_date)
            if to_date:
                mask = mask & (date_str_series <= to_date)
            combined_items = combined_items[mask].reset_index(drop=True)
            logger.info(f"Date filtered order items ({from_date} to {to_date}): {len(combined_items)} items remain.")

    # Re-build summary_df across all combined items
    if combined_items.empty:
        summary_df = pd.DataFrame(columns=["amazon-order-id", "total_qty", "sku", "order_status", "line_count"])
    else:
        summary_records = []
        for order_id, group in combined_items.groupby("amazon_order_id", sort=False):
            total_qty = int(group["quantity"].sum())
            distinct_skus = sorted({s for s in group["sku"].dropna() if s.strip()})
            joined_sku = " | ".join(distinct_skus) if distinct_skus else ""
            distinct_statuses = sorted({st for st in group["order_status"].dropna() if st.strip()})
            if len(distinct_statuses) > 1:
                final_status = "Mixed"
            elif len(distinct_statuses) == 1:
                final_status = distinct_statuses[0]
            else:
                final_status = "UNKNOWN"
            summary_records.append({
                "amazon-order-id": order_id,
                "total_qty": total_qty,
                "sku": joined_sku,
                "order_status": final_status,
                "line_count": len(group),
            })
        summary_df = pd.DataFrame(summary_records)

    return OrderParseResult(
        file_name="COMBINED_ORDERS",
        file_hash="COMBINED",
        items_df=combined_items,
        summary_df=summary_df,
        exceptions=all_exceptions,
    )


def load_multiple_payment_reports(
    files: List[Path],
    config: AppConfig,
    from_date: Optional[str] = None,
    to_date: Optional[str] = None,
    run_id: str = "MULTI_PMT",
) -> PaymentParseResult:
    """Load, combine, and deduplicate multiple payment reports idempotently."""
    if not files:
        empty_lines = pd.DataFrame()
        empty_pivot = pd.DataFrame(columns=["order_id", "total_net", "line_count", "order_amount", "shipping_fee", "refund_amount", "net_ex_shipping"])
        empty_non_order = pd.DataFrame()
        return PaymentParseResult("NONE", "", empty_lines, empty_pivot, empty_non_order, Decimal("0.00"), [])

    all_lines: List[pd.DataFrame] = []
    all_exceptions: List[PaymentException] = []
    seen_file_hashes: set[str] = set()

    for f in files:
        raw_copy = copy_to_raw_untouched(f, config.payments_raw_dir)
        parse_res = load_payment_report(raw_copy, config, run_id=run_id)
        if parse_res.file_hash in seen_file_hashes:
            logger.info(f"Skipping duplicate payment file with hash {parse_res.file_hash}: {f.name}")
            continue
        seen_file_hashes.add(parse_res.file_hash)
        if not parse_res.lines_df.empty:
            all_lines.append(parse_res.lines_df)
        all_exceptions.extend(parse_res.exceptions)

    if not all_lines:
        empty_pivot = pd.DataFrame(columns=["order_id", "total_net", "line_count", "order_amount", "shipping_fee", "refund_amount", "net_ex_shipping"])
        return PaymentParseResult("COMBINED", "", pd.DataFrame(), empty_pivot, pd.DataFrame(), Decimal("0.00"), all_exceptions)

    combined_lines = pd.concat(all_lines, ignore_index=True)

    # Deduplicate on row_hash
    if "row_hash" in combined_lines.columns and not combined_lines.empty:
        initial_len = len(combined_lines)
        combined_lines = combined_lines.drop_duplicates(subset=["row_hash"], keep="first")
        dedup_count = initial_len - len(combined_lines)
        if dedup_count > 0:
            logger.info(f"Deduplicated {dedup_count} overlapping payment lines across input files.")

    # Date filter on date_time_ist
    if (from_date or to_date) and not combined_lines.empty and "date_time_ist" in combined_lines.columns:
        valid_dates = combined_lines["date_time_ist"].dropna()
        if not valid_dates.empty:
            date_str_series = valid_dates.astype(str).str[:10]
            mask = pd.Series(True, index=combined_lines.index)
            if from_date:
                mask = mask & (date_str_series >= from_date)
            if to_date:
                mask = mask & (date_str_series <= to_date)
            combined_lines = combined_lines[mask].reset_index(drop=True)
            logger.info(f"Date filtered payment lines ({from_date} to {to_date}): {len(combined_lines)} lines remain.")

    grand_total = combined_lines["total"].sum() if not combined_lines.empty else Decimal("0.00")

    # Re-build order pivot & non-order df
    order_lines = combined_lines[combined_lines["clean_order_id"] != ""] if not combined_lines.empty else pd.DataFrame()
    non_order_lines = combined_lines[combined_lines["clean_order_id"] == ""] if not combined_lines.empty else pd.DataFrame()

    pivot_records = []
    if not order_lines.empty:
        for order_id, group in order_lines.groupby("clean_order_id", sort=False):
            total_net = group["total"].sum()
            line_count = len(group)
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

            net_ex_shipping = total_net - shipping_fee

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
    else:
        order_pivot_df = pd.DataFrame(columns=["order_id", "total_net", "line_count", "order_amount", "shipping_fee", "refund_amount", "net_ex_shipping"])

    # Non-order pivot
    if not non_order_lines.empty:
        non_order_records = []
        for (settlement_id, tx_type), group in non_order_lines.groupby(["settlement_id", "type"], sort=False):
            non_order_records.append({
                "settlement_id": settlement_id,
                "type": tx_type,
                "total_net": group["total"].sum(),
                "line_count": len(group),
                "description": group["description"].iloc[0] if "description" in group.columns else "",
            })
        non_order_df = pd.DataFrame(non_order_records)
    else:
        non_order_df = pd.DataFrame(columns=["settlement_id", "type", "total_net", "line_count", "description"])

    return PaymentParseResult(
        file_name="COMBINED_PAYMENTS",
        file_hash="COMBINED",
        lines_df=combined_lines,
        order_pivot_df=order_pivot_df,
        non_order_df=non_order_df,
        grand_total=grand_total,
        exceptions=all_exceptions,
    )


def run_reconciliation_pipeline(
    order_files: List[Path | str],
    payment_files: List[Path | str],
    config: AppConfig,
    auto_mode: bool = False,
    from_date: Optional[str] = None,
    to_date: Optional[str] = None,
    basis_override: Optional[str] = None,
    master_file: Optional[Path | str] = None,
    run_id: Optional[str] = None,
    output_excel_path: Optional[Path | str] = None,
) -> PipelineResult:
    """Run full reconciliation pipeline from input files to Excel report."""
    run_id = run_id or f"RUN_{datetime.now().strftime('%Y%m%d_%H%M%S')}"
    config.ensure_directories()

    # Determine delivery status basis
    if basis_override:
        if basis_override not in {"net_ex_shipping", "total_net"}:
            raise ValueError(f"Invalid basis '{basis_override}'. Must be 'net_ex_shipping' or 'total_net'.")
        config.delivery_status_basis = basis_override

    basis_used = config.delivery_status_basis

    all_inputs = [Path(p) for p in (order_files + payment_files)]
    inbox_processed_moves: List[Tuple[Path, Path]] = []

    # If auto mode, scan inbox
    if auto_mode:
        inbox_dir = config.inbox_dir
        if inbox_dir.exists():
            inbox_files = [
                f for f in inbox_dir.iterdir()
                if f.is_file() and f.suffix.lower() in {".txt", ".tsv", ".csv"}
            ]
            all_inputs.extend(inbox_files)

    # Classify files
    classified_orders, classified_payments, unrecognized = classify_files(all_inputs)

    if not classified_orders and not classified_payments:
        raise ValueError("No valid Amazon Order or Payment report files found to process.")

    logger.info(
        f"Pipeline start [{run_id}]: {len(classified_orders)} order report(s), "
        f"{len(classified_payments)} payment report(s), basis='{basis_used}'"
    )

    # Load and combine order reports
    order_res = load_multiple_order_reports(
        classified_orders, config, from_date=from_date, to_date=to_date, run_id=run_id
    )

    # Load and combine payment reports
    pmt_res = load_multiple_payment_reports(
        classified_payments, config, from_date=from_date, to_date=to_date, run_id=run_id
    )

    # Load master payout sheet (auto-discover if not provided)
    discovered_master = master_file
    if not discovered_master:
        for u in unrecognized:
            if "master" in u.name.lower() or "payout" in u.name.lower():
                discovered_master = u
                logger.info(f"Auto-detected master payout sheet from inputs: {u.name}")
                break

    from src.master import load_master_payout
    master_map = load_master_payout(discovered_master, config=config)

    # Execute reconciliation join
    recon_res = reconcile_orders_and_payments(
        order_res=order_res,
        payment_res=pmt_res,
        config=config,
        basis=basis_used,
        master_map=master_map,
        run_id=run_id,
    )

    # Generate Excel Report
    report_path = generate_reconciliation_excel_report(
        recon_res=recon_res,
        config=config,
        lines_df=pmt_res.lines_df,
        from_date=from_date,
        to_date=to_date,
        output_path=output_excel_path,
    )

    # In auto mode: move processed files from inbox to /data/processed/
    if auto_mode and config.inbox_dir.exists():
        timestamp_str = datetime.now().strftime("%Y%m%d_%H%M%S")
        for f in all_inputs:
            if f.parent.resolve() == config.inbox_dir.resolve() and f.exists():
                dest_file = config.processed_dir / f"{timestamp_str}_{f.name}"
                shutil.move(str(f), str(dest_file))
                logger.info(f"Moved processed inbox file: {f.name} -> {dest_file}")

    # Compute execution summary metrics
    recon_df = recon_res.recon_df
    status_counts: Dict[str, int] = {}
    status_amounts: Dict[str, Decimal] = {}
    for st in ["Delivered", "RVP", "RTO", "Not settled", "Cancelled"]:
        sub = recon_df[recon_df["delivery_status"] == st]
        status_counts[st] = len(sub)
        status_amounts[st] = sub["total_net"].apply(lambda v: v if isinstance(v, Decimal) else Decimal(str(v or 0))).sum()

    exception_severities = {"ERROR": 0, "WARNING": 0, "INFO": 0}
    for exc in recon_res.exceptions:
        sev = getattr(exc, "severity", "WARNING").upper()
        if sev in exception_severities:
            exception_severities[sev] += 1
        else:
            exception_severities[sev] = 1

    # Closing Amazon Balance Roll-forward
    non_order_lines = pmt_res.lines_df[pmt_res.lines_df["clean_order_id"] == ""] if not pmt_res.lines_df.empty else pd.DataFrame()
    transfers_total = non_order_lines[non_order_lines["type"].str.lower() == "transfer"]["total"].sum() if not non_order_lines.empty else Decimal("0.00")
    closing_balance = recon_res.grand_total - transfers_total

    result = PipelineResult(
        order_files=classified_orders,
        payment_files=classified_payments,
        total_order_items=len(order_res.items_df),
        unique_orders=len(order_res.summary_df),
        total_payment_lines=len(pmt_res.lines_df),
        settled_orders=len(pmt_res.order_pivot_df),
        non_order_lines=len(non_order_lines),
        grand_total=recon_res.grand_total,
        closing_balance=closing_balance,
        delivery_status_counts=status_counts,
        delivery_status_amounts=status_amounts,
        exceptions_count=len(recon_res.exceptions),
        exception_severities=exception_severities,
        report_path=report_path,
        basis_used=basis_used,
        from_date=from_date,
        to_date=to_date,
        excluded_payment_count=recon_res.excluded_payment_count,
        recon_res=recon_res,
    )

    return result


def format_console_summary(res: PipelineResult) -> str:
    """Format a clean, professional ASCII console summary dashboard."""
    date_range_str = f"{res.from_date} to {res.to_date}" if (res.from_date and res.to_date) else ("All available dates")
    shipped_orders = (
        res.delivery_status_counts.get("Delivered", 0)
        + res.delivery_status_counts.get("RVP", 0)
        + res.delivery_status_counts.get("RTO", 0)
    )

    order_files_str = "\n".join([f"    • {f.name}" for f in res.order_files]) if res.order_files else "    (None)"
    pmt_files_str = "\n".join([f"    • {f.name}" for f in res.payment_files]) if res.payment_files else "    (None)"

    lines = [
        "=" * 78,
        "          AMAZON.IN ORDER & PAYMENT RECONCILIATION DASHBOARD",
        "=" * 78,
        f"  Period Filter:    {date_range_str}",
        f"  Status Basis:     {res.basis_used}",
        f"  Order Reports ({len(res.order_files)}):",
        order_files_str,
        f"    -> Items Parsed: {res.total_order_items} | Unique Orders: {res.unique_orders}",
        f"  Payment Reports ({len(res.payment_files)}):",
        pmt_files_str,
        f"    -> Lines Parsed: {res.total_payment_lines} | Settled Orders: {res.settled_orders} | Non-Order: {res.non_order_lines}",
        f"    -> Payment Orders Excluded (Not in Order Master): {res.excluded_payment_count}",
        "-" * 78,
        "  DELIVERY STATUS BREAKDOWN:",
        f"    • Delivered:    {res.delivery_status_counts.get('Delivered', 0):>4} orders | Net: ₹{res.delivery_status_amounts.get('Delivered', Decimal('0.00')):>10.2f}",
        f"    • RVP (Return): {res.delivery_status_counts.get('RVP', 0):>4} orders | Net: ₹{res.delivery_status_amounts.get('RVP', Decimal('0.00')):>10.2f}",
        f"    • RTO (Return): {res.delivery_status_counts.get('RTO', 0):>4} orders | Net: ₹{res.delivery_status_amounts.get('RTO', Decimal('0.00')):>10.2f}",
        f"    • Not Settled:  {res.delivery_status_counts.get('Not settled', 0):>4} orders | Net: ₹{res.delivery_status_amounts.get('Not settled', Decimal('0.00')):>10.2f}",
        f"    • Cancelled:    {res.delivery_status_counts.get('Cancelled', 0):>4} orders | Net: ₹{res.delivery_status_amounts.get('Cancelled', Decimal('0.00')):>10.2f}",
        "-" * 78,
        "  FINANCIAL RECONCILIATION ROLL-FORWARD:",
        f"    • Shipped Orders Total:       {shipped_orders} orders",
        f"    • Grand Total (All Lines):    ₹{res.grand_total:>10.2f}",
        f"    • Closing Amazon Balance:     ₹{res.closing_balance:>10.2f} (Total less Bank Transfers)",
        "-" * 78,
        f"  DATA QUALITY & EXCEPTIONS:     {res.exceptions_count} audit records",
        f"    • Errors:   {res.exception_severities.get('ERROR', 0)}",
        f"    • Warnings: {res.exception_severities.get('WARNING', 0)}",
        f"    • Info:     {res.exception_severities.get('INFO', 0)}",
        "-" * 78,
        f"  EXCEL WORKBOOK SAVED TO:",
        f"    file:///{str(res.report_path).replace(chr(92), '/')}",
        "=" * 78,
    ]
    return "\n".join(lines)

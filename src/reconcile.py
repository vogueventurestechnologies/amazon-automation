"""Reconciliation engine joining order reports and payment reports for Amazon.in.

Implements Phase 4:
- LEFT JOIN from Order Summary (Master list) to Payment Pivot on order id.
- Output contains ONLY orders present in the order report (one row per amazon-order-id).
- Payment-report orders not in the order report are excluded from the main output table;
  logged to console and run log (e.g. "9 payment orders not in order report, excluded").
  Stored in payment_only_df for optional "Payment Only Orders" sheet if configured.
- Output columns:
    amazon-order-id, sku, quantity, order_status, total, net_ex_shipping, delivery_status, flags
- Determines delivery_status based on configured basis ('net_ex_shipping' default or 'total_net'):
    - order_status = Cancelled                    -> 'Cancelled'
    - no payment lines found                      -> 'Not settled' + exception SHIPPED_NOT_SETTLED
    - negative value                              -> 'RVP'
    - zero                                        -> 'RTO'
    - positive value                              -> 'Delivered'
- Computes flags (also recorded in exceptions):
    - STATUS_MISMATCH (info): order report says "Shipped" but delivery_status is RVP or RTO
    - REPLACEMENT_ORDER_COST: Order line total 0 with a shipping fee
    - REFUND_LOSS: refunded order with negative net
    - MIXED_STATUS: order has multiple item lines with differing statuses
    - ZERO_PRICE_SHIPPED: shipped item had 0.00 price
"""

from dataclasses import dataclass, field
from decimal import Decimal
import logging
from typing import Any, Dict, List, Optional, Set, Tuple
import pandas as pd

from src.config import AppConfig
from src.master import MasterSkuInfo, lookup_sku
from src.orders import OrderException, OrderParseResult
from src.payments import PaymentException, PaymentParseResult

logger = logging.getLogger("amazon_recon.reconcile")


@dataclass
class ReconException:
    """Unified exception record across orders, payments, and reconciliation."""

    run_id: str
    source: str
    reference_id: Optional[str]
    reason_code: str
    severity: str
    details: str


@dataclass
class ReconciliationResult:
    """Result of reconciling order report and payment report."""

    recon_df: pd.DataFrame
    status_counts: Dict[str, int]
    exceptions: List[ReconException]
    basis_used: str
    non_order_df: pd.DataFrame
    grand_total: Decimal
    payment_only_df: pd.DataFrame = field(default_factory=pd.DataFrame)
    excluded_payment_count: int = 0


def reconcile_orders_and_payments(
    order_res: OrderParseResult,
    payment_res: PaymentParseResult,
    config: AppConfig,
    basis: Optional[str] = None,
    master_map: Optional[Dict[str, MasterSkuInfo]] = None,
    run_id: str = "DEFAULT_RECON",
) -> ReconciliationResult:
    """Perform LEFT JOIN from order summary to payment pivot, assigning delivery status and flags.

    Args:
        order_res: OrderParseResult from Phase 2 (Master list).
        payment_res: PaymentParseResult from Phase 3.
        config: Application configuration.
        basis: Optional delivery status basis ('net_ex_shipping' or 'total_net'). Defaults to config.
        master_map: Optional mapping of SKU Name -> MasterSkuInfo for STYLE NAME & PAYOUT lookups.
        run_id: Current execution identifier.

    Returns:
        ReconciliationResult instance.
    """
    effective_basis = (basis or config.delivery_status_basis).strip().lower()
    if effective_basis not in {"net_ex_shipping", "total_net"}:
        raise ValueError(f"Invalid delivery_status_basis '{effective_basis}'. Must be 'net_ex_shipping' or 'total_net'")

    logger.info(f"Reconciling orders and payments using delivery status basis: '{effective_basis}' (LEFT JOIN: Order Master)")

    orders_df = order_res.summary_df.copy()
    payments_df = payment_res.order_pivot_df.copy()

    # Build lookup dictionaries
    orders_map: Dict[str, Dict[str, Any]] = {}
    order_id_sequence: List[str] = []
    if not orders_df.empty:
        order_id_sequence = list(orders_df["amazon-order-id"])
        orders_map = orders_df.set_index("amazon-order-id").to_dict("index")

    payments_map: Dict[str, Dict[str, Any]] = {}
    if not payments_df.empty:
        payments_map = payments_df.set_index("order_id").to_dict("index")

    # Identify payment orders not in the order report
    payment_only_ids = sorted(set(payments_map.keys()) - set(orders_map.keys()))
    excluded_payment_count = len(payment_only_ids)
    if excluded_payment_count > 0:
        logger.info(f"{excluded_payment_count} payment orders not in order report, excluded from main output.")

    # Build payment_only_df for separate sheet if needed
    payment_only_records = []
    for p_id in payment_only_ids:
        p_row = payments_map[p_id]
        payment_only_records.append({
            "order_id": p_id,
            "sku": p_row.get("sku", ""),
            "quantity": p_row.get("quantity", 0),
            "total_net": p_row.get("total_net", Decimal("0.00")),
            "net_ex_shipping": p_row.get("net_ex_shipping", Decimal("0.00")),
            "settlement_ids": p_row.get("settlement_ids", ""),
        })
    payment_only_df = pd.DataFrame(payment_only_records)

    # Build set of orders that had ZERO_PRICE_SHIPPED or MIXED_STATUS in order ingestion
    zero_price_order_items = {
        e.reference_id for e in order_res.exceptions if e.reason_code == "ZERO_PRICE_SHIPPED"
    }
    zero_price_order_ids = set()
    if not order_res.items_df.empty:
        zp_sub = order_res.items_df[order_res.items_df["order_item_id"].isin(zero_price_order_items)]
        zero_price_order_ids = set(zp_sub["amazon_order_id"].dropna().unique())

    mixed_status_order_ids = {
        e.reference_id for e in order_res.exceptions if e.reason_code == "MIXED_STATUS"
    }

    all_exceptions: List[ReconException] = []

    # Copy forward existing exceptions from orders and payments
    for e in order_res.exceptions:
        all_exceptions.append(ReconException(
            run_id=e.run_id,
            source=e.source,
            reference_id=e.reference_id,
            reason_code=e.reason_code,
            severity=e.severity,
            details=e.details,
        ))

    for pe in payment_res.exceptions:
        all_exceptions.append(ReconException(
            run_id=pe.run_id,
            source=pe.source,
            reference_id=pe.reference_id,
            reason_code=pe.reason_code,
            severity=pe.severity,
            details=pe.details,
        ))

    output_rows: List[Dict[str, Any]] = []

    # Iterate strictly over orders present in the order report (LEFT JOIN)
    for order_id in order_id_sequence:
        ord_info = orders_map[order_id]
        sku = ord_info.get("sku", "")
        quantity = ord_info.get("total_qty", 0)
        order_status = ord_info.get("order_status", "")

        sku_name = sku
        style_name = ""
        payout = None
        if master_map:
            m_info = lookup_sku(sku, master_map)
            style_name = m_info.style_name
            payout = m_info.payout

        in_payments = order_id in payments_map
        flags_list: List[str] = []

        if in_payments:
            pay_info = payments_map[order_id]
            total_net = pay_info.get("total_net", Decimal("0.00"))
            net_ex_shipping = pay_info.get("net_ex_shipping", Decimal("0.00"))
            order_amt = pay_info.get("order_amount", Decimal("0.00"))
            ship_fee = pay_info.get("shipping_fee", Decimal("0.00"))
            ref_amt = pay_info.get("refund_amount", Decimal("0.00"))
        else:
            pay_info = {}
            total_net = None
            net_ex_shipping = None
            order_amt = Decimal("0.00")
            ship_fee = Decimal("0.00")
            ref_amt = Decimal("0.00")

        # -------------------------------------------------------------------
        # Delivery Status Determination Rules (Section 8)
        # -------------------------------------------------------------------
        if order_status.strip().lower() == "cancelled":
            delivery_status = "Cancelled"
        elif not in_payments:
            delivery_status = "Not settled"
            flags_list.append("SHIPPED_NOT_SETTLED")
            all_exceptions.append(ReconException(
                run_id=run_id,
                source="reconciliation",
                reference_id=order_id,
                reason_code="SHIPPED_NOT_SETTLED",
                severity="INFO",
                details=f"Order {order_id} marked '{order_status}' in order report, but has no payment settlement lines",
            ))
        else:
            # Round to 2 decimals first before comparison
            basis_val = net_ex_shipping if effective_basis == "net_ex_shipping" else total_net
            rounded_val = basis_val.quantize(Decimal("0.01"))

            if rounded_val < Decimal("0.00"):
                delivery_status = "RVP"
            elif rounded_val == Decimal("0.00"):
                delivery_status = "RTO"
            else:
                delivery_status = "Delivered"

        # -------------------------------------------------------------------
        # Flag Rules (Section 8)
        # -------------------------------------------------------------------
        # REPLACEMENT_ORDER_COST: Order total 0 with a shipping fee
        if in_payments and order_amt == Decimal("0.00") and ship_fee < Decimal("0.00"):
            flags_list.append("REPLACEMENT_ORDER_COST")
            all_exceptions.append(ReconException(
                run_id=run_id,
                source="reconciliation",
                reference_id=order_id,
                reason_code="REPLACEMENT_ORDER_COST",
                severity="INFO",
                details=f"Replacement order {order_id} has 0 product sales but incurred shipping fee of {ship_fee}",
            ))

        # ZERO_PRICE_SHIPPED (from order report)
        if order_id in zero_price_order_ids:
            flags_list.append("ZERO_PRICE_SHIPPED")

        # STATUS_MISMATCH (info): order report says "Shipped" but delivery_status is RVP or RTO
        if order_status.lower().startswith("shipped") and delivery_status in ("RVP", "RTO"):
            flags_list.append("STATUS_MISMATCH")
            all_exceptions.append(ReconException(
                run_id=run_id,
                source="reconciliation",
                reference_id=order_id,
                reason_code="STATUS_MISMATCH",
                severity="INFO",
                details=f"Order {order_id} marked '{order_status}' in order report, but financial outcome is {delivery_status}",
            ))

        # REFUND_LOSS: refunded order with negative net
        if in_payments and ref_amt != Decimal("0.00") and total_net < Decimal("0.00"):
            flags_list.append("REFUND_LOSS")
            all_exceptions.append(ReconException(
                run_id=run_id,
                source="reconciliation",
                reference_id=order_id,
                reason_code="REFUND_LOSS",
                severity="WARNING",
                details=f"Refunded order {order_id} resulted in net loss of {total_net}",
            ))

        # MIXED_STATUS (from Phase 2)
        if order_id in mixed_status_order_ids:
            flags_list.append("MIXED_STATUS")

        # Deduplicate flags while preserving order
        clean_flags = []
        for f in flags_list:
            if f not in clean_flags:
                clean_flags.append(f)
        flags_str = ", ".join(clean_flags) if clean_flags else ""

        # Quantities for Delivered and Return (RVP)
        if delivery_status == "Delivered":
            delivered_qty = quantity if quantity > 0 else 1
            return_qty = 0
            prod_unit = payout if payout is not None else Decimal("0.00")
            net_production_cost = prod_unit * delivered_qty
        elif delivery_status in ("RVP", "RTO"):
            delivered_qty = 0
            return_qty = quantity if quantity > 0 else 1
            net_production_cost = Decimal("0.00")
        else:
            delivered_qty = 0
            return_qty = 0
            net_production_cost = Decimal("0.00")

        final_net = (total_net - net_production_cost) if total_net is not None else None

        output_rows.append({
            "amazon-order-id": order_id,
            "purchase_date": ord_info.get("purchase_date", ""),
            "product_name": ord_info.get("product_name", ""),
            "sku": sku,
            "sku_name": sku_name,
            "style_name": style_name,
            "payout": payout,
            "net_production_cost": net_production_cost,
            "production_cost": net_production_cost,
            "final_net": final_net,
            "quantity": quantity,
            "delivered_qty": delivered_qty,
            "return_qty": return_qty,
            "item_price": ord_info.get("item_price", Decimal("0.00")),
            "item_tax": ord_info.get("item_tax", Decimal("0.00")),
            "shipping_price": ord_info.get("shipping_price", Decimal("0.00")),
            "ship_city": ord_info.get("ship_city", ""),
            "ship_state": ord_info.get("ship_state", ""),
            "ship_postal_code": ord_info.get("ship_postal_code", ""),
            "order_status": order_status,
            "total": total_net,
            "total_net": total_net,
            "net_ex_shipping": net_ex_shipping,
            "delivery_status": delivery_status,
            "flags": flags_str,
        })

    recon_df = pd.DataFrame(output_rows)

    # Status counts calculated strictly on the master order report rows
    status_counts = recon_df["delivery_status"].value_counts().to_dict()

    logger.info(
        f"Reconciliation complete: {len(recon_df)} orders in output (LEFT JOIN). "
        f"Status counts: {status_counts}. Excluded payment orders: {excluded_payment_count}"
    )

    return ReconciliationResult(
        recon_df=recon_df,
        status_counts=status_counts,
        exceptions=all_exceptions,
        basis_used=effective_basis,
        non_order_df=payment_res.non_order_df,
        grand_total=payment_res.grand_total,
        payment_only_df=payment_only_df,
        excluded_payment_count=excluded_payment_count,
    )

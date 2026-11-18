"""FastAPI Web Application & Vercel Serverless Entrypoint for Amazon Reconciliation."""

from datetime import datetime
from decimal import Decimal
import os
from pathlib import Path
import shutil
import sys
import tempfile
from typing import Any, Dict, List, Optional
import uuid

# Ensure project root is in sys.path
PROJECT_ROOT = Path(__file__).resolve().parent.parent
if str(PROJECT_ROOT) not in sys.path:
    sys.path.insert(0, str(PROJECT_ROOT))

from fastapi import FastAPI, File, Form, HTTPException, UploadFile
from fastapi.middleware.cors import CORSMiddleware
from fastapi.responses import FileResponse, HTMLResponse, JSONResponse, Response

from src.config import AppConfig, load_config
from src.pipeline import PipelineResult, classify_files, format_console_summary, run_reconciliation_pipeline
from src.reporting.excel_export import _extract_design_and_size

app = FastAPI(
    title="Amazon.in Reconciliation Web App",
    description="Vercel & Local Web Application for Amazon.in Order + Payment Reconciliation",
    version="1.0.0",
)

app.add_middleware(
    CORSMiddleware,
    allow_origins=["*"],
    allow_credentials=True,
    allow_methods=["*"],
    allow_headers=["*"],
)

REPORTS_CACHE: Dict[str, Any] = {}
SAMPLE_ORDER_FILE = (
    PROJECT_ROOT / "data" / "raw" / "orders" / "157818151187020770.txt"
    if (PROJECT_ROOT / "data" / "raw" / "orders" / "157818151187020770.txt").exists()
    else Path("D:/2026/NOVEMBER MONTH/13-11-2026/157818151187020770.txt")
)
SAMPLE_PMT_FILE = (
    PROJECT_ROOT / "data" / "raw" / "payments" / "2026Aug1-2026Oct31CustomUnifiedTransaction.csv"
    if (PROJECT_ROOT / "data" / "raw" / "payments" / "2026Aug1-2026Oct31CustomUnifiedTransaction.csv").exists()
    else Path("D:/2026/NOVEMBER MONTH/13-11-2026/2026Aug1-2026Oct31CustomUnifiedTransaction.csv")
)
SAMPLE_MASTER_FILE = (
    PROJECT_ROOT / "data" / "master" / "MASTER PAYOUT SHEET.xlsx"
    if (PROJECT_ROOT / "data" / "master" / "MASTER PAYOUT SHEET.xlsx").exists()
    else Path("D:/2026/NOVEMBER MONTH/13-11-2026/MASTER PAYOUT SHEET.xlsx")
)


def _to_f(val: Any) -> float:
    """Safely convert any numeric / Decimal value to float."""
    if val is None:
        return 0.0
    try:
        return float(Decimal(str(val)))
    except Exception:
        return 0.0


def _build_response_payload(pipeline_res: PipelineResult, report_id: str, is_demo: bool = False) -> Dict[str, Any]:
    """Build fast, clean JSON payload directly from in-memory pipeline result."""
    recon_res = pipeline_res.recon_res
    recon_df = recon_res.recon_df if recon_res is not None else None

    orders_list = []
    sku_list = []
    exc_list = []

    if recon_df is not None and not recon_df.empty:
        for _, row in recon_df.iterrows():
            orders_list.append({
                "amazon_order_id": str(row.get("amazon-order-id", "")),
                "purchase_date": str(row.get("purchase_date", "") or ""),
                "product_name": str(row.get("product_name", "") or ""),
                "sku": str(row.get("sku", "") or ""),
                "sku_name": str(row.get("sku_name") or row.get("sku") or ""),
                "style_name": str(row.get("style_name", "") or ""),
                "payout": _to_f(row.get("payout")) if row.get("payout") is not None and str(row.get("payout")).strip() != "" else None,
                "net_production_cost": _to_f(row.get("net_production_cost", row.get("production_cost", 0.0))),
                "quantity": int(row.get("quantity", 0) or 0),
                "delivered_qty": int(row.get("delivered_qty", 0) or 0),
                "return_qty": int(row.get("return_qty", 0) or 0),
                "item_price": _to_f(row.get("item_price", 0.0)),
                "item_tax": _to_f(row.get("item_tax", 0.0)),
                "shipping_price": _to_f(row.get("shipping_price", 0.0)),
                "ship_city": str(row.get("ship_city", "") or ""),
                "ship_state": str(row.get("ship_state", "") or ""),
                "ship_postal_code": str(row.get("ship_postal_code", "") or ""),
                "order_status": str(row.get("order_status", "") or ""),
                "total_net": _to_f(row.get("total_net", 0.0)),
                "amazon_net": _to_f(row.get("amazon_net", 0.0)),
                "net_ex_shipping": _to_f(row.get("net_ex_shipping", 0.0)),
                "delivery_status": str(row.get("delivery_status", "") or ""),
                "flags": str(row.get("flags", "") or ""),
            })

        # SKU performance grouping
        for sku_val, group in recon_df.groupby("sku", sort=False):
            if not sku_val:
                continue
            design, size = _extract_design_and_size(sku_val)
            sold = int(group["quantity"].sum())
            delivered = len(group[group["delivery_status"] == "Delivered"])
            rvp = len(group[group["delivery_status"] == "RVP"])
            rto = len(group[group["delivery_status"] == "RTO"])
            shipped_c = delivered + rvp + rto
            ret_rate = round(((rvp + rto) / shipped_c) * 100, 1) if shipped_c > 0 else 0.0
            net_p = sum(_to_f(v) for v in group["total_net"])
            sku_list.append({
                "sku": sku_val,
                "design": design,
                "size": size,
                "total_orders": len(group),
                "qty_sold": sold,
                "delivered": delivered,
                "rvp": rvp,
                "rto": rto,
                "return_rate": ret_rate,
                "net_proceeds": net_p,
            })

    if recon_res is not None and recon_res.exceptions:
        for exc in recon_res.exceptions:
            exc_list.append({
                "run_id": str(getattr(exc, "run_id", "")),
                "source": str(getattr(exc, "source", "")),
                "reference_id": str(getattr(exc, "reference_id", "")),
                "reason_code": str(getattr(exc, "reason_code", "")),
                "severity": str(getattr(exc, "severity", "WARNING")),
                "details": str(getattr(exc, "details", "")),
            })

    report_name = (
        getattr(pipeline_res, "report_filename", None)
        or (pipeline_res.report_path.name if pipeline_res.report_path else "order_payment_recon.xlsx")
    )

    return {
        "success": True,
        "is_demo": is_demo,
        "report_id": report_id,
        "report_filename": report_name,
        "basis": pipeline_res.basis_used,
        "summary": {
            "delivered_count": pipeline_res.delivery_status_counts.get("Delivered", 0),
            "delivered_amount": _to_f(pipeline_res.delivery_status_amounts.get("Delivered", Decimal("0.00"))),
            "rvp_count": pipeline_res.delivery_status_counts.get("RVP", 0),
            "rvp_amount": _to_f(pipeline_res.delivery_status_amounts.get("RVP", Decimal("0.00"))),
            "rto_count": pipeline_res.delivery_status_counts.get("RTO", 0),
            "rto_amount": _to_f(pipeline_res.delivery_status_amounts.get("RTO", Decimal("0.00"))),
            "not_settled_count": pipeline_res.delivery_status_counts.get("Not settled", 0),
            "cancelled_count": pipeline_res.delivery_status_counts.get("Cancelled", 0),
            "grand_total": _to_f(pipeline_res.grand_total),
            "closing_balance": _to_f(pipeline_res.closing_balance),
            "total_shipped": (
                pipeline_res.delivery_status_counts.get("Delivered", 0)
                + pipeline_res.delivery_status_counts.get("RVP", 0)
                + pipeline_res.delivery_status_counts.get("RTO", 0)
            ),
            "exceptions_count": pipeline_res.exceptions_count,
        },
        "orders": orders_list,
        "sku_performance": sku_list,
        "exceptions": exc_list,
    }


@app.get("/", response_class=HTMLResponse)
def serve_dashboard():
    """Serve the modern interactive HTML5 + Tailwind CSS web dashboard."""
    html_file = PROJECT_ROOT / "static" / "index.html"
    if html_file.exists():
        return HTMLResponse(content=html_file.read_text(encoding="utf-8"))
    return HTMLResponse("<h2>Amazon.in Reconciliation Web App</h2><p>Frontend file static/index.html not found.</p>")


@app.get("/api/health")
def health_check():
    """Health check endpoint for Vercel / local monitoring."""
    return {
        "status": "healthy",
        "timestamp": datetime.now().isoformat(),
        "project_root": str(PROJECT_ROOT),
    }


@app.post("/api/reconcile")
def reconcile_files(
    order_file: Optional[UploadFile] = File(None),
    payment_file: Optional[UploadFile] = File(None),
    master_file: Optional[UploadFile] = File(None),
    basis: str = Form("total_net"),
    from_date: Optional[str] = Form(None),
    to_date: Optional[str] = Form(None),
):
    """Reconcile uploaded Order and Payment reports and return dashboard data + report ID."""
    has_order = isinstance(order_file, UploadFile) and bool(order_file.filename)
    has_pmt = isinstance(payment_file, UploadFile) and bool(payment_file.filename)
    has_master = isinstance(master_file, UploadFile) and bool(master_file.filename)

    if not has_order and not has_pmt and not has_master:
        raise HTTPException(status_code=400, detail="Please upload at least one Order or Payment report file.")

    temp_dir = Path(tempfile.mkdtemp(prefix="recon_upload_"))
    uploaded_paths: List[Path] = []
    master_path: Optional[Path] = None

    try:
        if has_order:
            ord_path = temp_dir / order_file.filename
            with open(ord_path, "wb") as f:
                content = order_file.file.read()
                f.write(content)
            uploaded_paths.append(ord_path)

        if has_pmt:
            pmt_path = temp_dir / payment_file.filename
            with open(pmt_path, "wb") as f:
                content = payment_file.file.read()
                f.write(content)
            uploaded_paths.append(pmt_path)

        if has_master:
            master_path = temp_dir / master_file.filename
            with open(master_path, "wb") as f:
                content = master_file.file.read()
                f.write(content)

        config = load_config()
        clean_from = from_date.strip() if isinstance(from_date, str) and from_date.strip() else None
        clean_to = to_date.strip() if isinstance(to_date, str) and to_date.strip() else None

        classified_orders, classified_pmts, unrec = classify_files(uploaded_paths)
        pipeline_res = run_reconciliation_pipeline(
            order_files=classified_orders,
            payment_files=classified_pmts,
            config=config,
            auto_mode=False,
            from_date=clean_from,
            to_date=clean_to,
            basis_override=basis,
            master_file=master_path,
            ephemeral=True,
            archive_raw=False,
        )

        report_id = str(uuid.uuid4())[:8]
        REPORTS_CACHE[report_id] = {
            "filename": pipeline_res.report_filename,
            "bytes": pipeline_res.report_bytes,
        }

        payload = _build_response_payload(pipeline_res, report_id, is_demo=False)
        return JSONResponse(content=payload)

    finally:
        try:
            shutil.rmtree(temp_dir, ignore_errors=True)
        except Exception:
            pass


@app.post("/api/demo")
def run_demo(basis: str = Form("total_net")):
    """Instantly run reconciliation on real samples for 1-click live demo."""
    if not (SAMPLE_ORDER_FILE.exists() and SAMPLE_PMT_FILE.exists()):
        raise HTTPException(status_code=404, detail="Demo sample files not found on server.")

    config = load_config()
    pipeline_res = run_reconciliation_pipeline(
        order_files=[SAMPLE_ORDER_FILE],
        payment_files=[SAMPLE_PMT_FILE],
        config=config,
        auto_mode=False,
        basis_override=basis,
        master_file=SAMPLE_MASTER_FILE if SAMPLE_MASTER_FILE.exists() else None,
        ephemeral=True,
        archive_raw=False,
    )

    report_id = str(uuid.uuid4())[:8]
    REPORTS_CACHE[report_id] = {
        "filename": pipeline_res.report_filename,
        "bytes": pipeline_res.report_bytes,
    }

    payload = _build_response_payload(pipeline_res, report_id, is_demo=True)
    return JSONResponse(content=payload)


@app.get("/api/download/{report_id}")
def download_excel(report_id: str):
    """Download the generated Excel reconciliation report directly from memory."""
    if report_id not in REPORTS_CACHE:
        raise HTTPException(status_code=404, detail="Report expired or not found. Please re-run reconciliation.")

    cached = REPORTS_CACHE[report_id]
    if isinstance(cached, dict) and "bytes" in cached and cached["bytes"]:
        filename = cached.get("filename", "order_payment_recon.xlsx")
        return Response(
            content=cached["bytes"],
            media_type="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            headers={"Content-Disposition": f'attachment; filename="{filename}"'},
        )
    elif isinstance(cached, Path) and cached.exists():
        return FileResponse(
            path=str(cached),
            filename=cached.name,
            media_type="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        )
    raise HTTPException(status_code=404, detail="Report file not found.")

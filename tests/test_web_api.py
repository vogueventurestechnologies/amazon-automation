"""Unit tests for FastAPI Web App and Vercel Serverless Endpoints."""

import json
from pathlib import Path
import pytest

from api.index import download_excel, health_check, run_demo, serve_dashboard

REAL_ORDER_FILE = Path("D:/2026/NOVEMBER MONTH/13-11-2026/157818151187020770.txt")
REAL_PAYMENT_FILE = Path("D:/2026/NOVEMBER MONTH/13-11-2026/2026Aug1-2026Oct31CustomUnifiedTransaction.csv")


def test_health_check_endpoint():
    """Verify /api/health returns healthy status."""
    res = health_check()
    assert res["status"] == "healthy"
    assert "timestamp" in res
    assert "project_root" in res


def test_serve_dashboard_html():
    """Verify root GET / serves the interactive Tailwind HTML5 dashboard."""
    res = serve_dashboard()
    assert res.status_code == 200
    html_content = res.body.decode("utf-8")
    assert "Amazon.in Reconciliation" in html_content
    assert "Order Report" in html_content
    assert "Payment Report" in html_content


def test_run_demo_endpoint_net_ex_shipping():
    """Verify /api/demo returns 200 and matches the expected 1 Delivered, 2 RVP, 1 RTO, 1 Cancelled breakdown."""
    if not (REAL_ORDER_FILE.exists() and REAL_PAYMENT_FILE.exists()):
        pytest.skip("Sample files not found on test machine")

    res = run_demo(basis="net_ex_shipping")
    assert res.status_code == 200
    data = json.loads(res.body.decode("utf-8"))

    assert data["success"] is True
    assert data["basis"] == "net_ex_shipping"
    assert data["summary"]["delivered_count"] == 1
    assert data["summary"]["rvp_count"] == 2
    assert data["summary"]["rto_count"] == 1
    assert data["summary"]["cancelled_count"] == 1
    assert data["summary"]["closing_balance"] == 310.46
    assert data["summary"]["grand_total"] == -755.26
    assert len(data["orders"]) == 5
    assert len(data["sku_performance"]) == 5

    # Test download endpoint for generated report
    report_id = data["report_id"]
    dl_res = download_excel(report_id)
    assert dl_res.status_code == 200
    assert "spreadsheetml.sheet" in dl_res.media_type


def test_run_demo_endpoint_total_net():
    """Verify /api/demo with total_net basis returns 3 RVP, 0 RTO, 1 Delivered, 1 Cancelled."""
    if not (REAL_ORDER_FILE.exists() and REAL_PAYMENT_FILE.exists()):
        pytest.skip("Sample files not found on test machine")

    res = run_demo(basis="total_net")
    assert res.status_code == 200
    data = json.loads(res.body.decode("utf-8"))

    assert data["basis"] == "total_net"
    assert data["summary"]["delivered_count"] == 1
    assert data["summary"]["rvp_count"] == 3
    assert data["summary"]["rto_count"] == 0
    assert data["summary"]["cancelled_count"] == 1


def test_reconcile_ephemeral_zero_disk_storage():
    """Verify customer uploads via /api/reconcile leave zero files on disk."""
    if not (REAL_ORDER_FILE.exists() and REAL_PAYMENT_FILE.exists()):
        pytest.skip("Sample files not found on test machine")

    import io
    from fastapi import UploadFile
    from api.index import reconcile_files

    proj_root = Path(__file__).resolve().parent.parent
    raw_orders_before = [f for f in (proj_root / "data" / "raw" / "orders").glob("*") if f.name != ".gitkeep"]
    raw_pmts_before = [f for f in (proj_root / "data" / "raw" / "payments").glob("*") if f.name != ".gitkeep"]
    reports_before = [f for f in (proj_root / "reports").glob("*.xlsx") if f.name != ".gitkeep"]

    with open(REAL_ORDER_FILE, "rb") as f_ord, open(REAL_PAYMENT_FILE, "rb") as f_pmt:
        ord_upload = UploadFile(filename=REAL_ORDER_FILE.name, file=io.BytesIO(f_ord.read()))
        pmt_upload = UploadFile(filename=REAL_PAYMENT_FILE.name, file=io.BytesIO(f_pmt.read()))

        res = reconcile_files(order_file=ord_upload, payment_file=pmt_upload, basis="total_net")
        assert res.status_code == 200
        data = json.loads(res.body.decode("utf-8"))
        assert data["success"] is True

        # Test download from memory
        dl_res = download_excel(data["report_id"])
        assert dl_res.status_code == 200
        assert len(dl_res.body) > 0

    # Verify zero files added to disk
    raw_orders_after = [f for f in (proj_root / "data" / "raw" / "orders").glob("*") if f.name != ".gitkeep"]
    raw_pmts_after = [f for f in (proj_root / "data" / "raw" / "payments").glob("*") if f.name != ".gitkeep"]
    reports_after = [f for f in (proj_root / "reports").glob("*.xlsx") if f.name != ".gitkeep"]

    assert len(raw_orders_after) == len(raw_orders_before) == 0
    assert len(raw_pmts_after) == len(raw_pmts_before) == 0
    assert len(reports_after) == len(reports_before) == 0


"""Unit and integration tests for Phase 6: CLI and Orchestration Pipeline.

Tests:
1. CLI no arguments: displays usage, returns 0.
2. CLI explicit files (--orders and --payments).
3. CLI auto mode (--auto): scans inbox, archives to raw, moves to processed.
4. CLI basis override: verifies --basis total_net vs net_ex_shipping.
5. Multi-file idempotent loading: deduplicates overlapping order items & payment lines.
6. Date filtering: validates --from and --to filtering on purchase and transaction dates.
"""

from decimal import Decimal
from pathlib import Path
import shutil
import openpyxl
import pytest

from main import main, parse_args
from src.config import AppConfig
from src.pipeline import classify_files, run_reconciliation_pipeline

REAL_ORDER_FILE = Path("D:/2026/NOVEMBER MONTH/13-11-2026/157818151187020770.txt")
REAL_PAYMENT_FILE = Path("D:/2026/NOVEMBER MONTH/13-11-2026/2026Aug1-2026Oct31CustomUnifiedTransaction.csv")


def test_cli_no_args_shows_usage(capsys):
    """Running CLI with no arguments exits cleanly with code 0 and usage advice."""
    code = main([])
    assert code == 0
    captured = capsys.readouterr()
    assert "No input files or --auto specified" in captured.out


def test_classify_files_auto_detection():
    """Verify header-based auto detection correctly classifies orders vs payments."""
    if not (REAL_ORDER_FILE.exists() and REAL_PAYMENT_FILE.exists()):
        pytest.skip("Real samples not found")

    # Pass files in reversed order
    orders, payments, unrec = classify_files([REAL_PAYMENT_FILE, REAL_ORDER_FILE])
    assert len(orders) == 1
    assert orders[0] == REAL_ORDER_FILE
    assert len(payments) == 1
    assert payments[0] == REAL_PAYMENT_FILE
    assert len(unrec) == 0


def test_cli_explicit_files_execution(tmp_project: AppConfig, capsys):
    """Test full CLI execution with explicit --orders and --payments arguments."""
    if not (REAL_ORDER_FILE.exists() and REAL_PAYMENT_FILE.exists()):
        pytest.skip("Real samples not found")

    cfg_path = tmp_project.project_root / "config" / "config.yaml"
    args = [
        "--orders", str(REAL_ORDER_FILE),
        "--payments", str(REAL_PAYMENT_FILE),
        "--config", str(cfg_path),
    ]

    code = main(args)
    assert code == 0
    captured = capsys.readouterr()
    assert "AMAZON.IN ORDER & PAYMENT RECONCILIATION DASHBOARD" in captured.out
    assert "Delivered:" in captured.out and "1 orders" in captured.out
    assert "RVP (Return):" in captured.out and "3 orders" in captured.out
    assert "RTO (Return):" in captured.out and "0 orders" in captured.out
    assert "Cancelled:" in captured.out and "1 orders" in captured.out
    assert "Payment Orders Excluded (Not in Order Master): 9" in captured.out

    # Verify report was generated
    report_file = tmp_project.reports_dir / "order_payment_recon_all.xlsx"
    assert report_file.exists()


def test_cli_auto_mode(tmp_project: AppConfig, capsys):
    """Test --auto mode: copies files from inbox to raw, reconciles, and moves inbox files to processed."""
    if not (REAL_ORDER_FILE.exists() and REAL_PAYMENT_FILE.exists()):
        pytest.skip("Real samples not found")

    inbox_order = tmp_project.inbox_dir / REAL_ORDER_FILE.name
    inbox_pmt = tmp_project.inbox_dir / REAL_PAYMENT_FILE.name

    shutil.copy2(REAL_ORDER_FILE, inbox_order)
    shutil.copy2(REAL_PAYMENT_FILE, inbox_pmt)

    cfg_path = tmp_project.project_root / "config" / "config.yaml"
    args = ["--auto", "--config", str(cfg_path)]

    code = main(args)
    assert code == 0

    # Verify inbox is now clean (files moved to processed)
    assert not inbox_order.exists()
    assert not inbox_pmt.exists()

    # Verify processed folder has the files
    processed_files = list(tmp_project.processed_dir.iterdir())
    assert len(processed_files) >= 2

    # Verify raw folder has untouched copies
    raw_orders = list(tmp_project.orders_raw_dir.iterdir())
    raw_pmts = list(tmp_project.payments_raw_dir.iterdir())
    assert len(raw_orders) >= 1
    assert len(raw_pmts) >= 1

    # Verify report exists
    report_file = tmp_project.reports_dir / "order_payment_recon_all.xlsx"
    assert report_file.exists()


def test_cli_basis_override_total_net(tmp_project: AppConfig, capsys):
    """Test CLI --basis total_net overrides default basis to produce 3 RVP, 0 RTO, 1 Delivered, 1 Cancelled."""
    if not (REAL_ORDER_FILE.exists() and REAL_PAYMENT_FILE.exists()):
        pytest.skip("Real samples not found")

    cfg_path = tmp_project.project_root / "config" / "config.yaml"
    args = [
        "--orders", str(REAL_ORDER_FILE),
        "--payments", str(REAL_PAYMENT_FILE),
        "--basis", "total_net",
        "--config", str(cfg_path),
    ]

    code = main(args)
    assert code == 0
    captured = capsys.readouterr()
    assert "Status Basis:     total_net" in captured.out
    assert "Delivered:" in captured.out and "1 orders" in captured.out
    assert "RVP (Return):" in captured.out and "3 orders" in captured.out
    assert "RTO (Return):" in captured.out and "0 orders" in captured.out
    assert "Cancelled:" in captured.out and "1 orders" in captured.out
    assert "Payment Orders Excluded (Not in Order Master): 9" in captured.out


def test_idempotent_multi_file_no_double_counting(tmp_project: AppConfig):
    """Passing the same file multiple times must deduplicate and not double-count amounts."""
    if not (REAL_ORDER_FILE.exists() and REAL_PAYMENT_FILE.exists()):
        pytest.skip("Real samples not found")

    # Pass payment file twice and order file twice
    res = run_reconciliation_pipeline(
        order_files=[REAL_ORDER_FILE, REAL_ORDER_FILE],
        payment_files=[REAL_PAYMENT_FILE, REAL_PAYMENT_FILE],
        config=tmp_project,
        run_id="TEST_IDEMPOTENT",
    )

    # Grand total across all rows must strictly remain -755.26, NOT doubled to -1510.52
    assert res.grand_total == Decimal("-755.26")
    assert res.unique_orders == 5
    assert res.settled_orders == 13
    assert res.delivery_status_counts["Delivered"] == 1
    assert res.delivery_status_counts["RVP"] == 3
    assert res.delivery_status_counts["RTO"] == 0
    assert res.delivery_status_counts["Cancelled"] == 1
    assert res.excluded_payment_count == 9


def test_date_range_filtering(tmp_project: AppConfig):
    """Test date filtering restricts the scope of reconciliation."""
    if not (REAL_ORDER_FILE.exists() and REAL_PAYMENT_FILE.exists()):
        pytest.skip("Real samples not found")

    # Filter to August 2026 only
    res = run_reconciliation_pipeline(
        order_files=[REAL_ORDER_FILE],
        payment_files=[REAL_PAYMENT_FILE],
        config=tmp_project,
        from_date="2026-08-01",
        to_date="2026-08-31",
        run_id="TEST_DATE_FILTER",
    )

    assert res.from_date == "2026-08-01"
    assert res.to_date == "2026-08-31"
    assert res.report_path.exists()
    assert "2026-08-01_2026-08-31" in res.report_path.name

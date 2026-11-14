"""Unit tests for Phase 1: Setup, Config, Logging, and CLI Skeleton."""

from decimal import Decimal
from pathlib import Path
import pytest

from main import main, parse_args
from src.config import AppConfig, load_config
from src.logger import setup_logger


def test_config_loading_and_directories(tmp_project: AppConfig):
    """Test that config loads correctly with all required Phase 1 parameters and folders."""
    cfg = tmp_project
    assert cfg.amount_tolerance == Decimal("0.01")
    assert cfg.timezone == "Asia/Kolkata"
    assert cfg.item_price_includes_tax is True
    assert cfg.delivery_status_basis == "total_net"
    assert cfg.include_payment_only_orders is False
    assert "Shipped - Delivered to Buyer" in cfg.order_status_map

    # Check directory structure
    assert cfg.orders_raw_dir.exists()
    assert cfg.payments_raw_dir.exists()
    assert cfg.inbox_dir.exists()
    assert cfg.processed_dir.exists()
    assert cfg.reports_dir.exists()
    assert cfg.logs_dir.exists()


def test_delivery_status_basis_validation(tmp_path: Path):
    """Test valid and invalid delivery_status_basis values."""
    raw_data_valid = {
        "delivery_status_basis": "total_net",
        "folder_paths": {},
    }
    cfg_valid = AppConfig(raw_config=raw_data_valid, project_root=tmp_path)
    assert cfg_valid.delivery_status_basis == "total_net"

    raw_data_invalid = {
        "delivery_status_basis": "invalid_basis",
        "folder_paths": {},
    }
    with pytest.raises(ValueError, match="Invalid delivery_status_basis"):
        AppConfig(raw_config=raw_data_invalid, project_root=tmp_path)


def test_logger_setup(tmp_path: Path):
    """Test logger writes output to file."""
    log_dir = tmp_path / "logs"
    logger = setup_logger(
        name="test_logger",
        log_dir=log_dir,
        log_file="test.log",
        level="INFO",
    )
    test_msg = "Phase 1 logger initialization test message"
    logger.info(test_msg)

    log_file = log_dir / "test.log"
    assert log_file.exists()
    content = log_file.read_text(encoding="utf-8")
    assert test_msg in content


def test_cli_argument_parsing():
    """Test CLI argument parsing for Phase 1 skeleton."""
    # Test --orders and --payments
    args = parse_args(["--orders", "order1.txt", "order2.tsv", "--payments", "pay1.csv"])
    assert args.orders == ["order1.txt", "order2.tsv"]
    assert args.payments == ["pay1.csv"]
    assert args.auto is False
    assert args.from_date is None
    assert args.to_date is None

    # Test --auto and date filters
    args_auto = parse_args(["--auto", "--from", "2026-08-01", "--to", "2026-08-31"])
    assert args_auto.auto is True
    assert args_auto.from_date == "2026-08-01"
    assert args_auto.to_date == "2026-08-31"


def test_main_execution_dry_run(tmp_project: AppConfig):
    """Test main execution with no args returns 0 cleanly."""
    cfg_file = tmp_project.project_root / "config" / "config.yaml"
    exit_code = main(["--config", str(cfg_file)])
    assert exit_code == 0

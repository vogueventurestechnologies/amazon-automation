"""Pytest fixtures for Amazon.in Order + Payment Reconciliation tests."""

from pathlib import Path
import pytest
import yaml

from src.config import AppConfig


@pytest.fixture
def tmp_project(tmp_path: Path) -> AppConfig:
    """Provide a temporary project root with isolated config and directories."""
    config_dir = tmp_path / "config"
    config_dir.mkdir(parents=True)
    config_file = config_dir / "config.yaml"

    config_data = {
        "amount_tolerance": 0.01,
        "timezone": "Asia/Kolkata",
        "item_price_includes_tax": True,
        "delivery_status_basis": "total_net",
        "include_payment_only_orders": False,
        "folder_paths": {
            "orders_raw": "data/raw/orders",
            "payments_raw": "data/raw/payments",
            "inbox": "data/inbox",
            "processed": "data/processed",
            "reports": "reports",
            "logs": "logs",
        },
        "report_output_dir": "reports",
        "order_status_map": {
            "Shipped - Delivered to Buyer": {"expected_receivable": True},
            "Shipped": {"expected_receivable": True},
            "Cancelled": {"expected_receivable": False},
        },
        "logging": {"level": "DEBUG", "log_file": "logs/test.log"},
    }

    with open(config_file, "w", encoding="utf-8") as f:
        yaml.dump(config_data, f)

    cfg = AppConfig(raw_config=config_data, project_root=tmp_path)
    cfg.ensure_directories()
    return cfg

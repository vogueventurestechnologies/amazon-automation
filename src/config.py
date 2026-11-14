"""Configuration loader and manager for Amazon.in Order + Payment Reconciliation."""

from dataclasses import dataclass, field
from decimal import Decimal
from pathlib import Path
from typing import Any, Dict, List, Optional
import yaml


VALID_DELIVERY_STATUS_BASES = {"net_ex_shipping", "total_net"}


@dataclass
class AppConfig:
    """Application configuration container."""

    raw_config: Dict[str, Any]
    project_root: Path

    # Resolved absolute paths
    orders_raw_dir: Path = field(init=False)
    payments_raw_dir: Path = field(init=False)
    inbox_dir: Path = field(init=False)
    processed_dir: Path = field(init=False)
    reports_dir: Path = field(init=False)
    logs_dir: Path = field(init=False)

    # Tolerances & Settings
    amount_tolerance: Decimal = field(init=False)
    timezone: str = field(init=False)
    item_price_includes_tax: bool = field(init=False)
    delivery_status_basis: str = field(init=False)
    include_payment_only_orders: bool = field(init=False)
    order_status_map: Dict[str, Dict[str, Any]] = field(init=False)
    report_output_dir: Path = field(init=False)

    def __post_init__(self) -> None:
        """Resolve paths relative to project root and validate settings."""
        folder_paths = self.raw_config.get("folder_paths", {})
        self.orders_raw_dir = (self.project_root / folder_paths.get("orders_raw", "data/raw/orders")).resolve()
        self.payments_raw_dir = (self.project_root / folder_paths.get("payments_raw", "data/raw/payments")).resolve()
        self.master_dir = (self.project_root / folder_paths.get("master", "data/master")).resolve()
        self.inbox_dir = (self.project_root / folder_paths.get("inbox", "data/inbox")).resolve()
        self.processed_dir = (self.project_root / folder_paths.get("processed", "data/processed")).resolve()
        
        rep_dir_str = self.raw_config.get("report_output_dir") or folder_paths.get("reports", "reports")
        self.reports_dir = (self.project_root / rep_dir_str).resolve()
        self.report_output_dir = self.reports_dir
        self.logs_dir = (self.project_root / folder_paths.get("logs", "logs")).resolve()

        # Tolerances & Settings
        self.amount_tolerance = Decimal(str(self.raw_config.get("amount_tolerance", "0.01")))
        self.timezone = str(self.raw_config.get("timezone", "Asia/Kolkata"))
        self.item_price_includes_tax = bool(self.raw_config.get("item_price_includes_tax", True))

        basis = str(self.raw_config.get("delivery_status_basis", "total_net")).strip().lower()
        if basis not in VALID_DELIVERY_STATUS_BASES:
            raise ValueError(
                f"Invalid delivery_status_basis '{basis}'. Must be one of: {VALID_DELIVERY_STATUS_BASES}"
            )
        self.delivery_status_basis = basis
        self.include_payment_only_orders = bool(self.raw_config.get("include_payment_only_orders", False))

        self.order_status_map = self.raw_config.get("order_status_map", {})

    def ensure_directories(self) -> None:
        """Ensure all required directories exist on disk."""
        self.orders_raw_dir.mkdir(parents=True, exist_ok=True)
        self.payments_raw_dir.mkdir(parents=True, exist_ok=True)
        self.master_dir.mkdir(parents=True, exist_ok=True)
        self.inbox_dir.mkdir(parents=True, exist_ok=True)
        self.processed_dir.mkdir(parents=True, exist_ok=True)
        self.reports_dir.mkdir(parents=True, exist_ok=True)
        self.logs_dir.mkdir(parents=True, exist_ok=True)


def load_config(config_path: str | Path | None = None) -> AppConfig:
    """Load configuration from a YAML file.

    Args:
        config_path: Optional path to config.yaml. Defaults to <project_root>/config/config.yaml.

    Returns:
        AppConfig instance.
    """
    if config_path is None:
        project_root = Path(__file__).resolve().parent.parent
        resolved_config_path = project_root / "config" / "config.yaml"
    else:
        resolved_config_path = Path(config_path).resolve()
        project_root = resolved_config_path.parent.parent

    if not resolved_config_path.exists():
        raise FileNotFoundError(f"Configuration file not found: {resolved_config_path}")

    with open(resolved_config_path, "r", encoding="utf-8") as f:
        data = yaml.safe_load(f) or {}

    config = AppConfig(raw_config=data, project_root=project_root)
    config.ensure_directories()
    return config

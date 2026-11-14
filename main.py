"""CLI entrypoint for Amazon.in Order + Payment Reconciliation Tool."""

import argparse
from datetime import datetime
import logging
import os
from pathlib import Path
import sys
from typing import List, Optional

from src.config import load_config
from src.logger import setup_logger
from src.pipeline import format_console_summary, run_reconciliation_pipeline

logger = logging.getLogger("amazon_recon")


def parse_args(args: Optional[List[str]] = None) -> argparse.Namespace:
    """Parse command line arguments for Amazon.in Order + Payment Reconciliation."""
    parser = argparse.ArgumentParser(
        description="Amazon.in Order + Payment Reconciliation Automation Tool"
    )
    parser.add_argument(
        "--orders",
        nargs="*",
        default=[],
        help="Path(s) to Amazon order report files (.txt / .tsv)",
    )
    parser.add_argument(
        "--payments",
        nargs="*",
        default=[],
        help="Path(s) to Amazon Custom Unified Transaction payment report files (.csv)",
    )
    parser.add_argument(
        "--master",
        type=str,
        default=None,
        help="Path to Master Payout Sheet (.xlsx / .csv) for SKU Name, Style Name & Payout lookups",
    )
    parser.add_argument(
        "--auto",
        action="store_true",
        help="Automatically detect and process order and payment files dropped in /data/inbox",
    )
    parser.add_argument(
        "--from",
        dest="from_date",
        type=str,
        default=None,
        help="Start date filter in IST (YYYY-MM-DD)",
    )
    parser.add_argument(
        "--to",
        dest="to_date",
        type=str,
        default=None,
        help="End date filter in IST (YYYY-MM-DD)",
    )
    parser.add_argument(
        "--basis",
        type=str,
        choices=["net_ex_shipping", "total_net"],
        default=None,
        help="Delivery status evaluation basis: net_ex_shipping (default) or total_net",
    )
    parser.add_argument(
        "--config",
        type=str,
        default=None,
        help="Path to config.yaml (defaults to config/config.yaml)",
    )
    parser.add_argument(
        "--open",
        action="store_true",
        help="Automatically open the generated Excel report upon completion",
    )
    return parser.parse_args(args)


def main(cli_args: Optional[List[str]] = None) -> int:
    """Main CLI execution flow."""
    args = parse_args(cli_args)
    config = load_config(args.config)

    log_settings = config.raw_config.get("logging", {})
    setup_logger(
        name="amazon_recon",
        log_dir=config.logs_dir,
        level=log_settings.get("level", "INFO"),
        log_file=log_settings.get("log_file", "logs/reconciliation.log").split("/")[-1],
        max_bytes=log_settings.get("max_bytes", 10485760),
        backup_count=log_settings.get("backup_count", 5),
    )

    logger.info("=" * 70)
    logger.info("Amazon.in Order + Payment Reconciliation Tool starting...")
    logger.info(f"Project root: {config.project_root}")
    logger.info(f"Delivery status basis: {args.basis or config.delivery_status_basis}")
    logger.info("=" * 70)

    if not (args.orders or args.payments or args.auto):
        print("\n[!] No input files or --auto specified.")
        print("Usage examples:")
        print("  python main.py --auto")
        print("  python main.py --orders path/to/orders.txt --payments path/to/payments.csv")
        print("  python main.py --auto --basis total_net --from 2026-08-01 --to 2026-10-31\n")
        return 0

    try:
        pipeline_res = run_reconciliation_pipeline(
            order_files=args.orders,
            payment_files=args.payments,
            config=config,
            auto_mode=args.auto,
            from_date=args.from_date,
            to_date=args.to_date,
            basis_override=args.basis,
            master_file=args.master,
        )

        dashboard_text = format_console_summary(pipeline_res)
        print("\n" + dashboard_text + "\n")

        if args.open and pipeline_res.report_path.exists():
            try:
                os.startfile(str(pipeline_res.report_path))
            except Exception as e:
                logger.warning(f"Could not open file automatically: {e}")

        return 0

    except Exception as e:
        logger.error(f"Reconciliation pipeline failed: {e}", exc_info=True)
        print(f"\n[ERROR] Pipeline failed: {e}\n")
        return 1


if __name__ == "__main__":
    sys.exit(main())

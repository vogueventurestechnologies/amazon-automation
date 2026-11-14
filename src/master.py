"""Master Payout Sheet Loader & SKU Lookup Engine.

Loads master payout sheet (Excel/CSV) containing:
- SKU Name (lookup key matching order report SKU)
- STYLE NAME
- PAYOUT (numeric rate)

Provides fast O(1) in-memory lookup with timestamped pickle caching.
"""

from dataclasses import dataclass
from decimal import Decimal, InvalidOperation
import logging
from pathlib import Path
import pickle
from typing import Any, Dict, Optional, Tuple
import pandas as pd

from src.config import AppConfig

logger = logging.getLogger("amazon_recon.master")


@dataclass
class MasterSkuInfo:
    """Master payout information for an SKU."""

    sku_name: str
    style_name: str
    payout: Optional[Decimal] = None


def _find_column_case_insensitive(columns: list, targets: list) -> Optional[str]:
    """Find the first matching column name from a list of candidates."""
    norm_cols = {str(c).strip().lower(): c for c in columns}
    for t in targets:
        if t.lower() in norm_cols:
            return norm_cols[t.lower()]
    # Substring search
    for norm_c, orig_c in norm_cols.items():
        for t in targets:
            if t.lower() in norm_c:
                return orig_c
    return None


def find_default_master_file(config: Optional[AppConfig] = None) -> Optional[Path]:
    """Auto-discover a master payout file from standard search locations."""
    search_dirs: list[Path] = []
    if config:
        if hasattr(config, "master_dir") and config.master_dir:
            search_dirs.append(config.master_dir)
        if hasattr(config, "inbox_dir") and config.inbox_dir:
            search_dirs.append(config.inbox_dir)
        search_dirs.append(config.project_root / "data" / "raw" / "master")

    # Local fallback
    cwd = Path.cwd()
    search_dirs.extend([
        cwd / "data" / "master",
        cwd / "data" / "raw" / "master",
        cwd / "data" / "inbox",
        Path("D:/2026/NOVEMBER MONTH/13-11-2026"),
    ])

    for s_dir in search_dirs:
        if not s_dir.exists():
            continue
        # Direct check for MASTER PAYOUT SHEET
        candidates = (
            list(s_dir.glob("*MASTER*PAYOUT*.xlsx"))
            + list(s_dir.glob("*master*payout*.xlsx"))
            + list(s_dir.glob("*payout*.xlsx"))
            + list(s_dir.glob("*master*.xlsx"))
            + list(s_dir.glob("*.xlsx"))
        )
        for cand in candidates:
            if not cand.name.startswith("~"):  # ignore temp excel lock files
                return cand.resolve()
    return None


def load_master_payout(
    master_path: Optional[str | Path] = None,
    config: Optional[AppConfig] = None,
    use_cache: bool = True,
) -> Dict[str, MasterSkuInfo]:
    """Load master payout sheet into an in-memory lookup dictionary keyed by normalized SKU.

    Args:
        master_path: Optional path to master sheet (.xlsx, .xls, .csv).
        config: Optional AppConfig instance.
        use_cache: Whether to use / write a .pkl cache file for sub-second loads.

    Returns:
        Dict[str, MasterSkuInfo] mapping SKU Name -> MasterSkuInfo.
    """
    target_file: Optional[Path] = None
    if master_path:
        p = Path(master_path).resolve()
        if p.exists():
            target_file = p
    if not target_file:
        target_file = find_default_master_file(config)

    if not target_file or not target_file.exists():
        logger.warning("No master payout sheet found. Style Name and Payout lookups will be empty.")
        return {}

    logger.info(f"Loading Master Payout Sheet: {target_file}")

    # Check cache
    cache_file = target_file.parent / f".cache_{target_file.stem}.pkl"
    if use_cache and cache_file.exists():
        try:
            target_mtime = target_file.stat().st_mtime
            cache_mtime = cache_file.stat().st_mtime
            if cache_mtime >= target_mtime:
                with open(cache_file, "rb") as f:
                    raw_cached = pickle.load(f)
                # Handle both legacy MasterSkuInfo and fast tuple dict
                if raw_cached:
                    sample_val = next(iter(raw_cached.values()))
                    if isinstance(sample_val, MasterSkuInfo):
                        return raw_cached
                    # Convert fast tuple (style, payout) to MasterSkuInfo mapping
                    fast_map = {
                        k: MasterSkuInfo(sku_name=k, style_name=v[0], payout=Decimal(v[1]) if v[1] else None)
                        for k, v in raw_cached.items()
                    }
                    logger.info(f"Loaded {len(fast_map)} master SKU records from fast cache in {cache_file.name}")
                    return fast_map
        except Exception as e:
            logger.warning(f"Could not load cache {cache_file}: {e}. Falling back to parsing file.")

    # Parse Excel / CSV
    try:
        if target_file.suffix.lower() in (".xlsx", ".xlsm", ".xls"):
            df = pd.read_excel(target_file)
        elif target_file.suffix.lower() in (".csv", ".txt"):
            df = pd.read_csv(target_file, dtype=str)
        else:
            logger.error(f"Unsupported master file format: {target_file.suffix}")
            return {}
    except Exception as e:
        logger.error(f"Failed to read master payout file {target_file}: {e}")
        return {}

    cols = list(df.columns)
    sku_col = _find_column_case_insensitive(cols, ["sku name", "sku_name", "sku", "item_sku", "product_sku"])
    style_col = _find_column_case_insensitive(cols, ["style name", "style_name", "style", "design"])
    payout_col = _find_column_case_insensitive(cols, ["payout", "pay out", "payout_amount", "rate"])

    if not sku_col:
        logger.warning(f"Could not find SKU column in master sheet {target_file.name}. Columns found: {cols}")
        return {}

    master_map: Dict[str, MasterSkuInfo] = {}
    cache_tuples: Dict[str, Tuple[str, str]] = {}

    # High-performance column zip: 10x faster than df.iterrows()
    sku_series = df[sku_col].fillna("").astype(str).str.strip()
    style_series = df[style_col].fillna("").astype(str).str.strip() if style_col else [""] * len(df)
    payout_series = df[payout_col] if payout_col else [None] * len(df)

    for raw_sku, style_val, raw_payout in zip(sku_series, style_series, payout_series):
        if not raw_sku or raw_sku in master_map:
            continue

        payout_val: Optional[Decimal] = None
        payout_str = ""
        if raw_payout is not None and not pd.isna(raw_payout):
            try:
                p_str = str(raw_payout).replace("₹", "").replace(",", "").strip()
                payout_val = Decimal(p_str)
                payout_str = str(payout_val)
            except (InvalidOperation, ValueError):
                pass

        master_map[raw_sku] = MasterSkuInfo(
            sku_name=raw_sku,
            style_name=style_val,
            payout=payout_val,
        )
        cache_tuples[raw_sku] = (style_val, payout_str)

    logger.info(f"Parsed {len(master_map)} unique master SKU lookup entries from {target_file.name}")

    # Write compact tuple cache for sub-second future loads
    if use_cache:
        try:
            with open(cache_file, "wb") as f:
                pickle.dump(cache_tuples, f, protocol=pickle.HIGHEST_PROTOCOL)
        except Exception as e:
            logger.warning(f"Failed to write master cache {cache_file}: {e}")

    return master_map


def lookup_sku(sku: str, master_map: Dict[str, MasterSkuInfo]) -> MasterSkuInfo:
    """Lookup an SKU in the master payout map. Returns default if not found."""
    clean_key = str(sku or "").strip()
    if clean_key in master_map:
        return master_map[clean_key]
    return MasterSkuInfo(sku_name=clean_key, style_name="", payout=None)

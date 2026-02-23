"""
Archive system for saving and loading monthly transaction data.
Stores categorized transactions as CSV files with JSON metadata
in a local archive folder for monthly accumulation and aggregation.
"""

import os
import json
import pandas as pd
from datetime import datetime
from pathlib import Path

ARCHIVE_DIR = Path(__file__).parent / "archives"


def ensure_archive_dir():
    """Create the archive directory if it doesn't exist."""
    ARCHIVE_DIR.mkdir(exist_ok=True)


def _sanitize_name(name: str) -> str:
    """Make a string safe for use in filenames."""
    return "".join(c if c.isalnum() or c in (" ", "-", "_") else "_" for c in name).strip()


def save_archive(
    df: pd.DataFrame,
    label: str,
    org_name: str = "",
    notes: str = "",
) -> str:
    """
    Save the current transaction data as an archived period.
    Returns the archive ID.
    """
    ensure_archive_dir()

    timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
    safe_label = _sanitize_name(label)
    archive_id = f"{safe_label}_{timestamp}"

    date_min = df["Date"].min()
    date_max = df["Date"].max()

    accounts = df["Account"].unique().tolist() if "Account" in df.columns else []

    metadata = {
        "id": archive_id,
        "label": label,
        "organization": org_name,
        "notes": notes,
        "created_at": datetime.now().isoformat(),
        "date_range_start": date_min.isoformat() if pd.notna(date_min) else None,
        "date_range_end": date_max.isoformat() if pd.notna(date_max) else None,
        "transaction_count": len(df),
        "accounts": accounts,
        "total_revenue": round(df[df["Amount"] > 0]["Amount"].sum(), 2),
        "total_expenses": round(abs(df[df["Amount"] < 0]["Amount"].sum()), 2),
    }

    csv_path = ARCHIVE_DIR / f"{archive_id}.csv"
    meta_path = ARCHIVE_DIR / f"{archive_id}.json"

    save_df = df.copy()
    save_df["Date"] = save_df["Date"].dt.strftime("%Y-%m-%d")
    save_df.to_csv(csv_path, index=False)

    with open(meta_path, "w") as f:
        json.dump(metadata, f, indent=2)

    return archive_id


def list_archives() -> list[dict]:
    """Return metadata for all saved archives, sorted newest first."""
    ensure_archive_dir()
    archives = []

    for meta_file in ARCHIVE_DIR.glob("*.json"):
        try:
            with open(meta_file) as f:
                meta = json.load(f)
            archives.append(meta)
        except (json.JSONDecodeError, KeyError):
            continue

    archives.sort(key=lambda x: x.get("created_at", ""), reverse=True)
    return archives


def load_archive(archive_id: str) -> tuple[pd.DataFrame, dict]:
    """Load a single archived period. Returns (dataframe, metadata)."""
    csv_path = ARCHIVE_DIR / f"{archive_id}.csv"
    meta_path = ARCHIVE_DIR / f"{archive_id}.json"

    if not csv_path.exists() or not meta_path.exists():
        raise FileNotFoundError(f"Archive '{archive_id}' not found.")

    with open(meta_path) as f:
        metadata = json.load(f)

    df = pd.read_csv(csv_path)
    df["Date"] = pd.to_datetime(df["Date"], format="mixed", errors="coerce")

    return df, metadata


def load_multiple_archives(archive_ids: list[str]) -> pd.DataFrame:
    """
    Load and combine multiple archived periods into a single dataframe.
    Deduplicates transactions that appear in overlapping periods.
    """
    all_dfs = []
    for aid in archive_ids:
        df, _ = load_archive(aid)
        df["_archive_id"] = aid
        all_dfs.append(df)

    if not all_dfs:
        return pd.DataFrame()

    combined = pd.concat(all_dfs, ignore_index=True)

    dedup_cols = ["Date", "Description", "Amount"]
    if "Account" in combined.columns:
        dedup_cols.append("Account")
    combined = combined.drop_duplicates(subset=dedup_cols, keep="first")
    combined = combined.drop(columns=["_archive_id"], errors="ignore")
    combined = combined.sort_values("Date").reset_index(drop=True)

    return combined


def delete_archive(archive_id: str) -> bool:
    """Delete an archived period. Returns True if successful."""
    csv_path = ARCHIVE_DIR / f"{archive_id}.csv"
    meta_path = ARCHIVE_DIR / f"{archive_id}.json"

    deleted = False
    if csv_path.exists():
        csv_path.unlink()
        deleted = True
    if meta_path.exists():
        meta_path.unlink()
        deleted = True

    return deleted


def get_archive_date_range(archives: list[dict]) -> tuple[str, str]:
    """Get the overall date range across multiple archive metadata entries."""
    starts = [a["date_range_start"] for a in archives if a.get("date_range_start")]
    ends = [a["date_range_end"] for a in archives if a.get("date_range_end")]

    if not starts or not ends:
        return "N/A", "N/A"

    return min(starts), max(ends)

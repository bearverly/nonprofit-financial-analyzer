"""
Bank statement parser that handles CSV and Excel files.
Auto-detects column mappings for date, description, and amount fields.
"""

import pandas as pd
import re
from typing import Optional


DATE_PATTERNS = [
    r"date", r"trans.*date", r"post.*date", r"effective.*date",
    r"settlement.*date", r"value.*date",
]

DESCRIPTION_PATTERNS = [
    r"desc", r"memo", r"narrat", r"detail", r"particular",
    r"reference", r"payee", r"transaction.*type", r"remark",
]

AMOUNT_PATTERNS = [r"amount", r"value", r"sum", r"total"]

DEBIT_PATTERNS = [r"debit", r"withdrawal", r"charge"]
CREDIT_PATTERNS = [r"credit", r"deposit", r"receipt"]


def _match_column(columns: list[str], patterns: list[str]) -> Optional[str]:
    """Find the first column name that matches any of the given patterns."""
    for col in columns:
        col_lower = col.lower().strip()
        for pattern in patterns:
            if re.search(pattern, col_lower):
                return col
    return None


def _parse_amount(value) -> float:
    """Parse a monetary value string to float."""
    if pd.isna(value):
        return 0.0
    if isinstance(value, (int, float)):
        return float(value)
    cleaned = re.sub(r"[,$\s]", "", str(value))
    cleaned = cleaned.replace("(", "-").replace(")", "")
    try:
        return float(cleaned)
    except ValueError:
        return 0.0


def detect_columns(df: pd.DataFrame) -> dict:
    """Auto-detect which columns map to date, description, amount, debit, credit."""
    columns = list(df.columns)

    mapping = {
        "date": _match_column(columns, DATE_PATTERNS),
        "description": _match_column(columns, DESCRIPTION_PATTERNS),
        "amount": _match_column(columns, AMOUNT_PATTERNS),
        "debit": _match_column(columns, DEBIT_PATTERNS),
        "credit": _match_column(columns, CREDIT_PATTERNS),
    }

    if not mapping["date"] and len(columns) > 0:
        for col in columns:
            try:
                pd.to_datetime(df[col].dropna().head(10), format="mixed")
                mapping["date"] = col
                break
            except (ValueError, TypeError):
                continue

    if not mapping["description"]:
        for col in columns:
            if col == mapping["date"]:
                continue
            if df[col].dtype == "object":
                sample = df[col].dropna().head(10)
                avg_len = sample.str.len().mean() if len(sample) > 0 else 0
                if avg_len > 5:
                    mapping["description"] = col
                    break

    return mapping


def parse_bank_statement(file, filename: str) -> tuple[pd.DataFrame, dict]:
    """
    Parse an uploaded bank statement file into a standardized DataFrame.
    Returns (dataframe, column_mapping).
    """
    ext = filename.lower().rsplit(".", 1)[-1] if "." in filename else ""

    if ext in ("xlsx", "xls"):
        df = pd.read_excel(file)
    else:
        try:
            df = pd.read_csv(file)
        except Exception:
            file.seek(0)
            df = pd.read_csv(file, encoding="latin-1")

    df.columns = [str(c).strip() for c in df.columns]

    if df.columns[0].startswith("Unnamed") or df.iloc[0].isna().sum() > len(df.columns) // 2:
        for i in range(min(5, len(df))):
            row = df.iloc[i]
            if row.notna().sum() >= 2:
                df.columns = [str(v).strip() if pd.notna(v) else f"Column_{j}" for j, v in enumerate(row)]
                df = df.iloc[i + 1:].reset_index(drop=True)
                break

    mapping = detect_columns(df)
    return df, mapping


def standardize_dataframe(df: pd.DataFrame, mapping: dict) -> pd.DataFrame:
    """
    Convert a raw bank statement DataFrame into a standardized format with
    columns: Date, Description, Amount.
    """
    result = pd.DataFrame()

    if mapping.get("date"):
        result["Date"] = pd.to_datetime(df[mapping["date"]], format="mixed", errors="coerce")
    else:
        result["Date"] = pd.NaT

    if mapping.get("description"):
        result["Description"] = df[mapping["description"]].astype(str).str.strip()
    else:
        result["Description"] = ""

    if mapping.get("amount"):
        result["Amount"] = df[mapping["amount"]].apply(_parse_amount)
    elif mapping.get("debit") and mapping.get("credit"):
        debits = df[mapping["debit"]].apply(_parse_amount)
        credits = df[mapping["credit"]].apply(_parse_amount)
        result["Amount"] = credits - debits
    elif mapping.get("debit"):
        result["Amount"] = -df[mapping["debit"]].apply(_parse_amount).abs()
    elif mapping.get("credit"):
        result["Amount"] = df[mapping["credit"]].apply(_parse_amount).abs()
    else:
        for col in df.columns:
            try:
                vals = df[col].apply(_parse_amount)
                if vals.abs().sum() > 0:
                    result["Amount"] = vals
                    break
            except Exception:
                continue
        if "Amount" not in result.columns:
            result["Amount"] = 0.0

    result = result.dropna(subset=["Date"]).reset_index(drop=True)
    result = result.sort_values("Date").reset_index(drop=True)

    return result

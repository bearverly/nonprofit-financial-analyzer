"""
Generates FASB-compliant nonprofit financial statements from categorized transactions.

Produces:
- Statement of Activities (ASC 958)
- Statement of Financial Position
- Statement of Functional Expenses
- Statement of Cash Flows
"""

import pandas as pd
from categorizer import (
    REVENUE_CATEGORIES, EXPENSE_CATEGORIES, FUNCTIONAL_CATEGORIES,
    get_functional_classification, get_category_type,
)


def statement_of_activities(df: pd.DataFrame, org_name: str = "Organization") -> dict:
    """
    Generate Statement of Activities (nonprofit income statement).
    Shows revenue, expenses, and change in net assets.
    """
    period_start = df["Date"].min().strftime("%B %d, %Y") if len(df) > 0 else "N/A"
    period_end = df["Date"].max().strftime("%B %d, %Y") if len(df) > 0 else "N/A"

    revenue_data = {}
    for cat in REVENUE_CATEGORIES:
        mask = df["Category"] == cat
        total = df.loc[mask, "Amount"].sum()
        if total != 0:
            revenue_data[cat] = round(total, 2)

    expense_data = {}
    for cat in EXPENSE_CATEGORIES:
        if cat == "Internal Account Transfer":
            continue
        mask = df["Category"] == cat
        total = abs(df.loc[mask, "Amount"].sum())
        if total != 0:
            expense_data[cat] = round(total, 2)

    transfer_in = df.loc[
        (df["Category"] == "Internal Account Transfer") & (df["Amount"] > 0), "Amount"
    ].sum()
    transfer_out = abs(df.loc[
        (df["Category"] == "Internal Account Transfer") & (df["Amount"] < 0), "Amount"
    ].sum())
    net_transfers = round(transfer_in - transfer_out, 2)

    total_revenue = round(sum(revenue_data.values()), 2)
    total_expenses = round(sum(expense_data.values()), 2)
    change_in_net_assets = round(total_revenue - total_expenses + net_transfers, 2)

    return {
        "title": f"Statement of Activities",
        "organization": org_name,
        "period": f"For the Period {period_start} to {period_end}",
        "without_donor_restrictions": {
            "revenue": revenue_data,
            "total_revenue": total_revenue,
            "expenses": expense_data,
            "total_expenses": total_expenses,
            "net_transfers": net_transfers,
            "change_in_net_assets": change_in_net_assets,
        },
    }


def statement_of_financial_position(
    df: pd.DataFrame,
    beginning_cash: float = 0.0,
    other_assets: float = 0.0,
    liabilities: float = 0.0,
    org_name: str = "Organization",
) -> dict:
    """
    Generate Statement of Financial Position (nonprofit balance sheet).
    """
    period_end = df["Date"].max().strftime("%B %d, %Y") if len(df) > 0 else "N/A"

    net_cash_activity = round(df["Amount"].sum(), 2)
    ending_cash = round(beginning_cash + net_cash_activity, 2)

    total_assets = round(ending_cash + other_assets, 2)
    net_assets_without_restriction = round(total_assets - liabilities, 2)
    total_liabilities_and_net_assets = round(liabilities + net_assets_without_restriction, 2)

    return {
        "title": "Statement of Financial Position",
        "organization": org_name,
        "as_of": f"As of {period_end}",
        "assets": {
            "Cash and Cash Equivalents": ending_cash,
            "Other Assets": other_assets,
            "Total Assets": total_assets,
        },
        "liabilities": {
            "Total Liabilities": liabilities,
        },
        "net_assets": {
            "Without Donor Restrictions": net_assets_without_restriction,
            "With Donor Restrictions": 0.0,
            "Total Net Assets": net_assets_without_restriction,
        },
        "total_liabilities_and_net_assets": total_liabilities_and_net_assets,
    }


def statement_of_functional_expenses(
    df: pd.DataFrame, org_name: str = "Organization"
) -> dict:
    """
    Generate Statement of Functional Expenses.
    Breaks expenses down by natural classification (rows) and functional classification (columns).
    """
    period_start = df["Date"].min().strftime("%B %d, %Y") if len(df) > 0 else "N/A"
    period_end = df["Date"].max().strftime("%B %d, %Y") if len(df) > 0 else "N/A"

    expense_df = df[df["Amount"] < 0].copy()
    expense_df["AbsAmount"] = expense_df["Amount"].abs()
    expense_df["Functional"] = expense_df["Category"].apply(get_functional_classification)

    natural_categories = [c for c in EXPENSE_CATEGORIES if c not in FUNCTIONAL_CATEGORIES]
    natural_categories = natural_categories + ["Other Expenses"]

    seen = set()
    unique_natural = []
    for c in natural_categories:
        if c not in seen:
            seen.add(c)
            unique_natural.append(c)

    table = {}
    for nat_cat in unique_natural:
        row = {}
        cat_df = expense_df[expense_df["Category"] == nat_cat]
        for func_cat in FUNCTIONAL_CATEGORIES:
            val = cat_df.loc[cat_df["Functional"] == func_cat, "AbsAmount"].sum()
            row[func_cat] = round(val, 2)
        row["Total"] = round(sum(row.values()), 2)
        if row["Total"] > 0:
            table[nat_cat] = row

    totals = {}
    for func_cat in FUNCTIONAL_CATEGORIES:
        totals[func_cat] = round(
            sum(table.get(nc, {}).get(func_cat, 0) for nc in unique_natural), 2
        )
    totals["Total"] = round(sum(totals.values()), 2)

    func_direct = {}
    for func_cat in FUNCTIONAL_CATEGORIES:
        direct = expense_df.loc[
            (expense_df["Category"] == func_cat), "AbsAmount"
        ].sum()
        if direct > 0:
            existing = table.get(func_cat, {func: 0 for func in FUNCTIONAL_CATEGORIES})
            existing[func_cat] = existing.get(func_cat, 0) + round(direct, 2)
            existing["Total"] = round(sum(v for k, v in existing.items() if k != "Total"), 2)
            table[func_cat + " (Direct)"] = existing
            totals[func_cat] = round(totals.get(func_cat, 0) + direct, 2)
            totals["Total"] = round(sum(v for k, v in totals.items() if k != "Total"), 2)

    return {
        "title": "Statement of Functional Expenses",
        "organization": org_name,
        "period": f"For the Period {period_start} to {period_end}",
        "functional_categories": FUNCTIONAL_CATEGORIES,
        "table": table,
        "totals": totals,
    }


def statement_of_cash_flows(
    df: pd.DataFrame,
    beginning_cash: float = 0.0,
    org_name: str = "Organization",
) -> dict:
    """
    Generate Statement of Cash Flows using the direct method.
    """
    period_start = df["Date"].min().strftime("%B %d, %Y") if len(df) > 0 else "N/A"
    period_end = df["Date"].max().strftime("%B %d, %Y") if len(df) > 0 else "N/A"

    operating_inflows = {}
    operating_outflows = {}

    for cat in REVENUE_CATEGORIES:
        total = df.loc[df["Category"] == cat, "Amount"].sum()
        if total > 0:
            operating_inflows[f"Cash received from {cat.lower()}"] = round(total, 2)

    for cat in EXPENSE_CATEGORIES:
        total = abs(df.loc[df["Category"] == cat, "Amount"].sum())
        if total > 0:
            operating_outflows[f"Cash paid for {cat.lower()}"] = round(-total, 2)

    total_inflows = round(sum(operating_inflows.values()), 2)
    total_outflows = round(sum(operating_outflows.values()), 2)
    net_operating = round(total_inflows + total_outflows, 2)

    ending_cash = round(beginning_cash + net_operating, 2)

    return {
        "title": "Statement of Cash Flows",
        "organization": org_name,
        "period": f"For the Period {period_start} to {period_end}",
        "operating_activities": {
            "inflows": operating_inflows,
            "outflows": operating_outflows,
            "net": net_operating,
        },
        "investing_activities": {"net": 0.0},
        "financing_activities": {"net": 0.0},
        "net_change_in_cash": net_operating,
        "beginning_cash": beginning_cash,
        "ending_cash": ending_cash,
    }


def generate_all_statements(
    df: pd.DataFrame,
    org_name: str = "Organization",
    beginning_cash: float = 0.0,
    other_assets: float = 0.0,
    liabilities: float = 0.0,
) -> dict:
    """Generate all four financial statements."""
    return {
        "activities": statement_of_activities(df, org_name),
        "position": statement_of_financial_position(
            df, beginning_cash, other_assets, liabilities, org_name
        ),
        "functional_expenses": statement_of_functional_expenses(df, org_name),
        "cash_flows": statement_of_cash_flows(df, beginning_cash, org_name),
    }

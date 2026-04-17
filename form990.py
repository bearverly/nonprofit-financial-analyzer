"""
IRS Form 990 preparation worksheet for 501(c)(3) nonprofit organizations.

Maps categorized transaction data to the key schedules of Form 990:
- Part I:    Summary
- Part VIII: Statement of Revenue
- Part IX:   Statement of Functional Expenses
- Part X:    Balance Sheet (Statement of Financial Position)
- Schedule A: Public Charity Status (support schedule)
"""

import pandas as pd
from categorizer import (
    REVENUE_CATEGORIES, EXPENSE_CATEGORIES, FUNCTIONAL_CATEGORIES,
    get_functional_classification, get_parent_category,
)


def _matching_categories(base_names: list[str], all_categories: list[str]) -> list[str]:
    """Return all categories (including sub-categories) that match the given base names."""
    result = []
    for cat in all_categories:
        parent = get_parent_category(cat)
        if cat in base_names or parent in base_names:
            result.append(cat)
    return result


def _sum_categories(df: pd.DataFrame, categories: list[str], positive_only=False, negative_only=False) -> float:
    all_cats = _matching_categories(categories, df["Category"].unique().tolist())
    mask = df["Category"].isin(all_cats)
    subset = df[mask]
    if positive_only:
        subset = subset[subset["Amount"] > 0]
    if negative_only:
        subset = subset[subset["Amount"] < 0]
    return round(abs(subset["Amount"].sum()), 2)


def generate_part1_summary(df: pd.DataFrame, beginning_net_assets: float = 0.0) -> dict:
    """Part I: Summary -- high-level revenue, expenses, net assets."""
    contributions = _sum_categories(df, ["Donations & Contributions", "Grants"], positive_only=True)
    program_revenue = _sum_categories(df, ["Program Service Revenue"], positive_only=True)
    investment_income = _sum_categories(df, ["Investment Income"], positive_only=True)
    other_revenue = _sum_categories(df, ["Fundraising Event Revenue", "Other Revenue"], positive_only=True)
    total_revenue = round(contributions + program_revenue + investment_income + other_revenue, 2)

    grants_paid = 0.0
    benefits_paid = 0.0
    salaries = _sum_categories(df, ["Salaries & Benefits"], negative_only=True)

    program_expenses = 0.0
    mgmt_expenses = 0.0
    fundraising_expenses = 0.0
    expense_df = df[(df["Amount"] < 0) & (df["Category"] != "Internal Account Transfer")].copy()
    expense_df["Functional"] = expense_df["Category"].apply(get_functional_classification)
    for _, row in expense_df.iterrows():
        amt = abs(row["Amount"])
        func = row["Functional"]
        if func == "Program Services":
            program_expenses += amt
        elif func == "Management & General":
            mgmt_expenses += amt
        elif func == "Fundraising":
            fundraising_expenses += amt

    total_expenses = round(program_expenses + mgmt_expenses + fundraising_expenses, 2)
    net_change = round(total_revenue - total_expenses, 2)
    ending_net_assets = round(beginning_net_assets + net_change, 2)

    return {
        "line_1": contributions,       # Contributions and grants
        "line_2": program_revenue,      # Program service revenue
        "line_3": investment_income,    # Investment income
        "line_4": other_revenue,        # Other revenue
        "line_5_to_9": 0.0,            # Gain/loss from asset sales, gaming, etc.
        "line_12": total_revenue,       # Total revenue
        "line_13": grants_paid,         # Grants and similar amounts paid
        "line_14": benefits_paid,       # Benefits paid to members
        "line_15": salaries,            # Salaries and compensation
        "line_16a": 0.0,                # Professional fundraising fees
        "line_16b": 0.0,               # Total fundraising expenses
        "line_17": total_expenses,      # Other expenses
        "line_18": total_expenses,      # Total expenses
        "line_19": net_change,          # Revenue less expenses
        "line_20": beginning_net_assets,
        "line_21": 0.0,                # Other changes
        "line_22": ending_net_assets,   # Ending net assets
    }


def generate_part8_revenue(df: pd.DataFrame) -> dict:
    """Part VIII: Statement of Revenue -- detailed revenue breakdown."""
    contributions_gifts = _sum_categories(df, ["Donations & Contributions"], positive_only=True)
    grants_govt = _sum_categories(df, ["Grants"], positive_only=True)
    total_contributions = round(contributions_gifts + grants_govt, 2)

    program_categories = {}
    prog_df = df[(df["Category"] == "Program Service Revenue") & (df["Amount"] > 0)]
    if len(prog_df) > 0:
        for desc_group, group_df in prog_df.groupby("Description"):
            short_desc = str(desc_group)[:50]
            program_categories[short_desc] = round(group_df["Amount"].sum(), 2)
    total_program = round(sum(program_categories.values()), 2)

    investment_income = _sum_categories(df, ["Investment Income"], positive_only=True)

    fundraising_gross = _sum_categories(df, ["Fundraising Event Revenue"], positive_only=True)
    fundraising_expenses = _sum_categories(df, ["Fundraising"], negative_only=True)
    fundraising_net = round(fundraising_gross - fundraising_expenses, 2)

    other_revenue = _sum_categories(df, ["Other Revenue"], positive_only=True)

    total_revenue = round(
        total_contributions + total_program + investment_income
        + fundraising_net + other_revenue, 2
    )

    return {
        "1a_federated_campaigns": 0.0,
        "1b_membership_dues": 0.0,
        "1c_fundraising_events": 0.0,
        "1d_related_organizations": 0.0,
        "1e_govt_grants": grants_govt,
        "1f_all_other_contributions": contributions_gifts,
        "1h_total_contributions": total_contributions,
        "2a_program_services": program_categories,
        "2a_total": total_program,
        "3_investment_income": investment_income,
        "4_income_from_investment_of_tax_exempt_bonds": 0.0,
        "5_royalties": 0.0,
        "6a_gross_rents": 0.0,
        "7a_gross_from_sales": 0.0,
        "8a_fundraising_events_gross": fundraising_gross,
        "8b_fundraising_expenses": fundraising_expenses,
        "8c_fundraising_net": fundraising_net,
        "9a_gaming_gross": 0.0,
        "10a_gross_sales_of_inventory": 0.0,
        "11_other_revenue": other_revenue,
        "12_total_revenue": total_revenue,
    }


def generate_part9_expenses(df: pd.DataFrame) -> dict:
    """
    Part IX: Statement of Functional Expenses.
    Each line broken into Total, Program, Management & General, and Fundraising.
    """
    expense_df = df[(df["Amount"] < 0) & (df["Category"] != "Internal Account Transfer")].copy()
    expense_df["AbsAmount"] = expense_df["Amount"].abs()
    expense_df["Functional"] = expense_df["Category"].apply(get_functional_classification)

    line_items = {
        "1_grants_to_domestic_orgs": [],
        "2_grants_to_domestic_individuals": [],
        "3_grants_to_foreign": [],
        "4_benefits_to_members": [],
        "5_compensation_current_officers": ["Salaries & Benefits"],
        "6_compensation_disqualified": [],
        "7_other_salaries": [],
        "8_pension_plans": [],
        "9_other_employee_benefits": [],
        "10_payroll_taxes": [],
        "11a_management_fees": ["Professional Services"],
        "11b_legal_fees": [],
        "11c_accounting_fees": [],
        "11d_lobbying_fees": [],
        "11e_professional_fundraising": [],
        "11f_investment_management": [],
        "11g_other_fees": [],
        "12_advertising": [],
        "13_office_expenses": ["Supplies & Materials"],
        "14_information_technology": ["Management & General"],
        "15_royalties": [],
        "16_occupancy": ["Facilities & Occupancy"],
        "17_travel": ["Travel & Transportation"],
        "18_third_party_payments": [],
        "19_other_expenses_a": ["Program Services", "League Equipment"],
        "19_other_expenses_b": ["Fundraising"],
        "19_other_expenses_c": ["Insurance"],
        "19_other_expenses_d": ["Other Expenses"],
    }

    result = {}
    for line_key, categories in line_items.items():
        row = {"Total": 0.0, "Program Services": 0.0, "Management & General": 0.0, "Fundraising": 0.0}
        for cat in categories:
            cat_df = expense_df[expense_df["Category"] == cat]
            for func in FUNCTIONAL_CATEGORIES:
                val = cat_df.loc[cat_df["Functional"] == func, "AbsAmount"].sum()
                row[func] = round(row[func] + val, 2)
        row["Total"] = round(row["Program Services"] + row["Management & General"] + row["Fundraising"], 2)
        result[line_key] = row

    totals = {"Total": 0.0, "Program Services": 0.0, "Management & General": 0.0, "Fundraising": 0.0}
    for row in result.values():
        for func in ["Total"] + FUNCTIONAL_CATEGORIES:
            totals[func] = round(totals[func] + row[func], 2)
    result["25_total"] = totals

    return result


def generate_part10_balance_sheet(
    df: pd.DataFrame,
    beginning_cash: float = 0.0,
    other_assets: float = 0.0,
    liabilities: float = 0.0,
    beginning_net_assets: float = 0.0,
) -> dict:
    """Part X: Balance Sheet -- assets, liabilities, net assets."""
    net_activity = round(df["Amount"].sum(), 2)
    ending_cash = round(beginning_cash + net_activity, 2)
    total_assets = round(ending_cash + other_assets, 2)

    non_transfer = df[df["Category"] != "Internal Account Transfer"]
    total_revenue = non_transfer[non_transfer["Amount"] > 0]["Amount"].sum()
    total_expenses = abs(non_transfer[non_transfer["Amount"] < 0]["Amount"].sum())
    net_change = round(total_revenue - total_expenses, 2)
    ending_net_assets = round(beginning_net_assets + net_change, 2)

    return {
        "assets": {
            "1_cash": ending_cash,
            "2_savings": 0.0,
            "3_pledges_receivable": 0.0,
            "4_accounts_receivable": 0.0,
            "5_loans_receivable": 0.0,
            "6_loans_receivable_other": 0.0,
            "7_inventories": 0.0,
            "8_prepaid_expenses": 0.0,
            "10a_land_buildings": 0.0,
            "10b_less_depreciation": 0.0,
            "10c_net_land_buildings": 0.0,
            "11_investments_public": 0.0,
            "12_investments_other": other_assets,
            "13_program_related_investments": 0.0,
            "14_intangible_assets": 0.0,
            "15_other_assets": 0.0,
            "16_total_assets": total_assets,
        },
        "liabilities": {
            "17_accounts_payable": 0.0,
            "18_grants_payable": 0.0,
            "19_deferred_revenue": 0.0,
            "20_tax_exempt_bond_liabilities": 0.0,
            "21_escrow_account": 0.0,
            "22_loans_from_officers": 0.0,
            "23_secured_mortgages": 0.0,
            "24_unsecured_notes": 0.0,
            "25_other_liabilities": liabilities,
            "26_total_liabilities": liabilities,
        },
        "net_assets": {
            "27_unrestricted": ending_net_assets,
            "28_temporarily_restricted": 0.0,
            "29_permanently_restricted": 0.0,
            "30_capital_stock": 0.0,
            "31_paid_in_capital": 0.0,
            "32_retained_earnings": 0.0,
            "33_total_net_assets": ending_net_assets,
            "34_total_liabilities_and_net_assets": round(liabilities + ending_net_assets, 2),
        },
    }


def generate_schedule_a_support(df: pd.DataFrame) -> dict:
    """
    Schedule A: Public Charity Status and Public Support Test.
    Summarizes support sources for the 33-1/3% public support test.
    """
    contributions = _sum_categories(df, ["Donations & Contributions"], positive_only=True)
    grants = _sum_categories(df, ["Grants"], positive_only=True)
    program_revenue = _sum_categories(df, ["Program Service Revenue"], positive_only=True)
    investment_income = _sum_categories(df, ["Investment Income"], positive_only=True)
    fundraising = _sum_categories(df, ["Fundraising Event Revenue"], positive_only=True)
    other = _sum_categories(df, ["Other Revenue"], positive_only=True)

    total_support = round(contributions + grants + program_revenue
                          + investment_income + fundraising + other, 2)

    public_support = round(contributions + grants + program_revenue + fundraising, 2)
    public_support_pct = round((public_support / total_support * 100) if total_support > 0 else 0, 1)

    return {
        "gifts_grants_contributions": contributions,
        "government_grants": grants,
        "program_service_revenue": program_revenue,
        "investment_income": investment_income,
        "fundraising_revenue": fundraising,
        "other_revenue": other,
        "total_support": total_support,
        "public_support": public_support,
        "public_support_percentage": public_support_pct,
        "meets_33_percent_test": public_support_pct >= 33.33,
    }


def generate_form990_data(
    df: pd.DataFrame,
    org_name: str = "Organization",
    ein: str = "",
    fiscal_year_start: str = "",
    fiscal_year_end: str = "",
    beginning_cash: float = 0.0,
    other_assets: float = 0.0,
    liabilities: float = 0.0,
    beginning_net_assets: float = 0.0,
) -> dict:
    """Generate all Form 990 worksheet data."""
    period_start = df["Date"].min().strftime("%m/%d/%Y") if len(df) > 0 else "N/A"
    period_end = df["Date"].max().strftime("%m/%d/%Y") if len(df) > 0 else "N/A"

    return {
        "header": {
            "organization_name": org_name,
            "ein": ein,
            "fiscal_year_start": fiscal_year_start or period_start,
            "fiscal_year_end": fiscal_year_end or period_end,
            "form_type": "990",
        },
        "part1_summary": generate_part1_summary(df, beginning_net_assets),
        "part8_revenue": generate_part8_revenue(df),
        "part9_expenses": generate_part9_expenses(df),
        "part10_balance_sheet": generate_part10_balance_sheet(
            df, beginning_cash, other_assets, liabilities, beginning_net_assets
        ),
        "schedule_a": generate_schedule_a_support(df),
    }

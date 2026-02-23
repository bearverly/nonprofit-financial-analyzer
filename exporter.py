"""
Export financial statements to Excel workbooks with formatted sheets.
"""

import io
import pandas as pd


def _write_header(ws, row, title, fmt):
    """Write a bold header row."""
    ws.write(row, 0, title, fmt)
    return row + 1


def _write_line_item(ws, row, label, amount, label_fmt, money_fmt, indent=0):
    """Write a single line item with label and amount."""
    prefix = "    " * indent
    ws.write(row, 0, f"{prefix}{label}", label_fmt)
    ws.write(row, 1, amount, money_fmt)
    return row + 1


def export_to_excel(statements: dict, transactions_df: pd.DataFrame) -> bytes:
    """Export all financial statements and transaction detail to an Excel workbook."""
    output = io.BytesIO()

    with pd.ExcelWriter(output, engine="xlsxwriter") as writer:
        workbook = writer.book

        title_fmt = workbook.add_format({
            "bold": True, "font_size": 14, "bottom": 2,
            "font_name": "Calibri",
        })
        subtitle_fmt = workbook.add_format({
            "italic": True, "font_size": 11, "font_name": "Calibri",
        })
        section_fmt = workbook.add_format({
            "bold": True, "font_size": 11, "bottom": 1,
            "font_name": "Calibri",
        })
        label_fmt = workbook.add_format({"font_name": "Calibri", "font_size": 11})
        money_fmt = workbook.add_format({
            "num_format": "#,##0.00", "font_name": "Calibri", "font_size": 11,
        })
        total_money_fmt = workbook.add_format({
            "num_format": "#,##0.00", "bold": True, "top": 1, "bottom": 6,
            "font_name": "Calibri", "font_size": 11,
        })
        total_label_fmt = workbook.add_format({
            "bold": True, "top": 1, "bottom": 6,
            "font_name": "Calibri", "font_size": 11,
        })
        bold_fmt = workbook.add_format({
            "bold": True, "font_name": "Calibri", "font_size": 11,
        })
        bold_money_fmt = workbook.add_format({
            "num_format": "#,##0.00", "bold": True,
            "font_name": "Calibri", "font_size": 11,
        })

        # --- Statement of Activities ---
        act = statements["activities"]
        ws = workbook.add_worksheet("Statement of Activities")
        ws.set_column(0, 0, 45)
        ws.set_column(1, 1, 18)
        row = 0
        row = _write_header(ws, row, act["title"], title_fmt)
        ws.write(row, 0, act["organization"], subtitle_fmt)
        row += 1
        ws.write(row, 0, act["period"], subtitle_fmt)
        row += 2

        data = act["without_donor_restrictions"]

        ws.write(row, 0, "REVENUE AND SUPPORT", section_fmt)
        ws.write(row, 1, "", section_fmt)
        row += 1
        for label, val in data["revenue"].items():
            row = _write_line_item(ws, row, label, val, label_fmt, money_fmt, indent=1)
        row = _write_line_item(ws, row, "Total Revenue and Support", data["total_revenue"],
                               bold_fmt, bold_money_fmt)
        row += 1

        ws.write(row, 0, "EXPENSES", section_fmt)
        ws.write(row, 1, "", section_fmt)
        row += 1
        for label, val in data["expenses"].items():
            row = _write_line_item(ws, row, label, val, label_fmt, money_fmt, indent=1)
        row = _write_line_item(ws, row, "Total Expenses", data["total_expenses"],
                               bold_fmt, bold_money_fmt)
        row += 1

        row = _write_line_item(ws, row, "CHANGE IN NET ASSETS", data["change_in_net_assets"],
                               total_label_fmt, total_money_fmt)

        # --- Statement of Financial Position ---
        pos = statements["position"]
        ws = workbook.add_worksheet("Financial Position")
        ws.set_column(0, 0, 45)
        ws.set_column(1, 1, 18)
        row = 0
        row = _write_header(ws, row, pos["title"], title_fmt)
        ws.write(row, 0, pos["organization"], subtitle_fmt)
        row += 1
        ws.write(row, 0, pos["as_of"], subtitle_fmt)
        row += 2

        ws.write(row, 0, "ASSETS", section_fmt)
        ws.write(row, 1, "", section_fmt)
        row += 1
        for label, val in pos["assets"].items():
            fmt_l = bold_fmt if "Total" in label else label_fmt
            fmt_m = bold_money_fmt if "Total" in label else money_fmt
            row = _write_line_item(ws, row, label, val, fmt_l, fmt_m,
                                   indent=0 if "Total" in label else 1)
        row += 1

        ws.write(row, 0, "LIABILITIES", section_fmt)
        ws.write(row, 1, "", section_fmt)
        row += 1
        for label, val in pos["liabilities"].items():
            row = _write_line_item(ws, row, label, val, label_fmt, money_fmt, indent=1)
        row += 1

        ws.write(row, 0, "NET ASSETS", section_fmt)
        ws.write(row, 1, "", section_fmt)
        row += 1
        for label, val in pos["net_assets"].items():
            fmt_l = bold_fmt if "Total" in label else label_fmt
            fmt_m = bold_money_fmt if "Total" in label else money_fmt
            row = _write_line_item(ws, row, label, val, fmt_l, fmt_m,
                                   indent=0 if "Total" in label else 1)
        row += 1
        row = _write_line_item(ws, row, "TOTAL LIABILITIES AND NET ASSETS",
                               pos["total_liabilities_and_net_assets"],
                               total_label_fmt, total_money_fmt)

        # --- Statement of Functional Expenses ---
        func = statements["functional_expenses"]
        ws = workbook.add_worksheet("Functional Expenses")
        ws.set_column(0, 0, 35)
        for i in range(1, 5):
            ws.set_column(i, i, 18)
        row = 0
        row = _write_header(ws, row, func["title"], title_fmt)
        ws.write(row, 0, func["organization"], subtitle_fmt)
        row += 1
        ws.write(row, 0, func["period"], subtitle_fmt)
        row += 2

        headers = [""] + func["functional_categories"] + ["Total"]
        for col, h in enumerate(headers):
            ws.write(row, col, h, section_fmt)
        row += 1

        for nat_cat, values in func["table"].items():
            ws.write(row, 0, nat_cat, label_fmt)
            for col, func_cat in enumerate(func["functional_categories"], 1):
                ws.write(row, col, values.get(func_cat, 0), money_fmt)
            ws.write(row, len(func["functional_categories"]) + 1, values.get("Total", 0), money_fmt)
            row += 1

        ws.write(row, 0, "Total Expenses", total_label_fmt)
        for col, func_cat in enumerate(func["functional_categories"], 1):
            ws.write(row, col, func["totals"].get(func_cat, 0), total_money_fmt)
        ws.write(row, len(func["functional_categories"]) + 1,
                 func["totals"].get("Total", 0), total_money_fmt)

        # --- Statement of Cash Flows ---
        cf = statements["cash_flows"]
        ws = workbook.add_worksheet("Cash Flows")
        ws.set_column(0, 0, 50)
        ws.set_column(1, 1, 18)
        row = 0
        row = _write_header(ws, row, cf["title"], title_fmt)
        ws.write(row, 0, cf["organization"], subtitle_fmt)
        row += 1
        ws.write(row, 0, cf["period"], subtitle_fmt)
        row += 2

        ws.write(row, 0, "OPERATING ACTIVITIES", section_fmt)
        ws.write(row, 1, "", section_fmt)
        row += 1

        for label, val in cf["operating_activities"]["inflows"].items():
            row = _write_line_item(ws, row, label, val, label_fmt, money_fmt, indent=1)
        for label, val in cf["operating_activities"]["outflows"].items():
            row = _write_line_item(ws, row, label, val, label_fmt, money_fmt, indent=1)

        row = _write_line_item(ws, row, "Net Cash from Operating Activities",
                               cf["operating_activities"]["net"], bold_fmt, bold_money_fmt)
        row += 1

        ws.write(row, 0, "INVESTING ACTIVITIES", section_fmt)
        ws.write(row, 1, "", section_fmt)
        row += 1
        row = _write_line_item(ws, row, "Net Cash from Investing Activities",
                               cf["investing_activities"]["net"], bold_fmt, bold_money_fmt)
        row += 1

        ws.write(row, 0, "FINANCING ACTIVITIES", section_fmt)
        ws.write(row, 1, "", section_fmt)
        row += 1
        row = _write_line_item(ws, row, "Net Cash from Financing Activities",
                               cf["financing_activities"]["net"], bold_fmt, bold_money_fmt)
        row += 1

        row = _write_line_item(ws, row, "Net Change in Cash", cf["net_change_in_cash"],
                               bold_fmt, bold_money_fmt)
        row = _write_line_item(ws, row, "Beginning Cash Balance", cf["beginning_cash"],
                               label_fmt, money_fmt)
        row = _write_line_item(ws, row, "ENDING CASH BALANCE", cf["ending_cash"],
                               total_label_fmt, total_money_fmt)

        # --- Transaction Detail ---
        export_df = transactions_df.copy()
        if "Date" in export_df.columns:
            export_df["Date"] = export_df["Date"].dt.strftime("%m/%d/%Y")

        has_account = "Account" in export_df.columns
        detail_cols = ["Date", "Description", "Amount", "Category", "Functional"]
        if has_account:
            detail_cols = ["Date", "Account", "Description", "Amount", "Category", "Functional"]
        available_cols = [c for c in detail_cols if c in export_df.columns]
        export_df[available_cols].to_excel(writer, sheet_name="Transaction Detail", index=False)

        detail_ws = writer.sheets["Transaction Detail"]
        if has_account:
            detail_ws.set_column(0, 0, 14)
            detail_ws.set_column(1, 1, 20)
            detail_ws.set_column(2, 2, 50)
            detail_ws.set_column(3, 3, 15)
            detail_ws.set_column(4, 4, 30)
            detail_ws.set_column(5, 5, 25)
        else:
            detail_ws.set_column(0, 0, 14)
            detail_ws.set_column(1, 1, 50)
            detail_ws.set_column(2, 2, 15)
            detail_ws.set_column(3, 3, 30)
            detail_ws.set_column(4, 4, 25)

    output.seek(0)
    return output.getvalue()


def export_form990_to_excel(data: dict, transactions_df: pd.DataFrame) -> bytes:
    """Export Form 990 preparation worksheet to a formatted Excel workbook."""
    output = io.BytesIO()

    with pd.ExcelWriter(output, engine="xlsxwriter") as writer:
        workbook = writer.book

        title_fmt = workbook.add_format({
            "bold": True, "font_size": 14, "bottom": 2, "font_name": "Calibri",
        })
        subtitle_fmt = workbook.add_format({
            "italic": True, "font_size": 11, "font_name": "Calibri",
        })
        section_fmt = workbook.add_format({
            "bold": True, "font_size": 11, "bottom": 1, "font_name": "Calibri",
            "bg_color": "#E8EAF6",
        })
        label_fmt = workbook.add_format({"font_name": "Calibri", "font_size": 11})
        money_fmt = workbook.add_format({
            "num_format": "#,##0.00", "font_name": "Calibri", "font_size": 11,
        })
        bold_fmt = workbook.add_format({
            "bold": True, "font_name": "Calibri", "font_size": 11,
        })
        bold_money_fmt = workbook.add_format({
            "num_format": "#,##0.00", "bold": True, "font_name": "Calibri", "font_size": 11,
        })
        total_label_fmt = workbook.add_format({
            "bold": True, "top": 1, "bottom": 6, "font_name": "Calibri", "font_size": 11,
        })
        total_money_fmt = workbook.add_format({
            "num_format": "#,##0.00", "bold": True, "top": 1, "bottom": 6,
            "font_name": "Calibri", "font_size": 11,
        })
        line_num_fmt = workbook.add_format({
            "font_name": "Calibri", "font_size": 10, "font_color": "#666666",
        })
        header_col_fmt = workbook.add_format({
            "bold": True, "font_size": 10, "bottom": 1, "font_name": "Calibri",
            "text_wrap": True, "align": "center",
        })

        header = data["header"]

        # --- Part I: Summary ---
        ws = workbook.add_worksheet("Part I - Summary")
        ws.set_column(0, 0, 8)
        ws.set_column(1, 1, 55)
        ws.set_column(2, 2, 18)
        row = 0
        ws.write(row, 0, "Form 990 — Part I: Summary", title_fmt)
        row += 1
        ws.write(row, 0, header["organization_name"], subtitle_fmt)
        ws.write(row, 2, f"EIN: {header['ein']}" if header["ein"] else "", subtitle_fmt)
        row += 1
        ws.write(row, 0, f"Tax Year: {header['fiscal_year_start']} to {header['fiscal_year_end']}", subtitle_fmt)
        row += 2

        p1 = data["part1_summary"]
        p1_lines = [
            ("1", "Contributions, gifts, grants, and similar amounts received", p1["line_1"]),
            ("2", "Program service revenue", p1["line_2"]),
            ("3", "Investment income", p1["line_3"]),
            ("4", "Other revenue", p1["line_4"]),
            ("12", "Total revenue (add lines 1 through 4)", p1["line_12"]),
            ("", "", None),
            ("13", "Grants and similar amounts paid", p1["line_13"]),
            ("15", "Salaries, other compensation, employee benefits", p1["line_15"]),
            ("18", "Total expenses", p1["line_18"]),
            ("", "", None),
            ("19", "Revenue less expenses (line 12 minus line 18)", p1["line_19"]),
            ("20", "Total net assets or fund balances — beginning of year", p1["line_20"]),
            ("22", "Total net assets or fund balances — end of year", p1["line_22"]),
        ]

        for line_num, desc, val in p1_lines:
            if val is None:
                row += 1
                continue
            is_total = line_num in ("12", "18", "19", "22")
            ws.write(row, 0, line_num, line_num_fmt)
            ws.write(row, 1, desc, bold_fmt if is_total else label_fmt)
            ws.write(row, 2, val, bold_money_fmt if is_total else money_fmt)
            row += 1

        # --- Part VIII: Revenue ---
        ws = workbook.add_worksheet("Part VIII - Revenue")
        ws.set_column(0, 0, 8)
        ws.set_column(1, 1, 55)
        ws.set_column(2, 2, 18)
        row = 0
        ws.write(row, 0, "Form 990 — Part VIII: Statement of Revenue", title_fmt)
        row += 2

        p8 = data["part8_revenue"]
        ws.write(row, 0, "", section_fmt)
        ws.write(row, 1, "Contributions, Gifts, Grants", section_fmt)
        ws.write(row, 2, "", section_fmt)
        row += 1
        ws.write(row, 0, "1e", line_num_fmt)
        ws.write(row, 1, "Government grants (contributions)", label_fmt)
        ws.write(row, 2, p8["1e_govt_grants"], money_fmt)
        row += 1
        ws.write(row, 0, "1f", line_num_fmt)
        ws.write(row, 1, "All other contributions, gifts, grants", label_fmt)
        ws.write(row, 2, p8["1f_all_other_contributions"], money_fmt)
        row += 1
        ws.write(row, 0, "1h", line_num_fmt)
        ws.write(row, 1, "Total (add lines 1a through 1f)", bold_fmt)
        ws.write(row, 2, p8["1h_total_contributions"], bold_money_fmt)
        row += 2

        ws.write(row, 0, "", section_fmt)
        ws.write(row, 1, "Program Service Revenue", section_fmt)
        ws.write(row, 2, "", section_fmt)
        row += 1
        for desc, val in p8["2a_program_services"].items():
            ws.write(row, 0, "2a", line_num_fmt)
            ws.write(row, 1, desc, label_fmt)
            ws.write(row, 2, val, money_fmt)
            row += 1
        ws.write(row, 0, "2f", line_num_fmt)
        ws.write(row, 1, "Total program service revenue", bold_fmt)
        ws.write(row, 2, p8["2a_total"], bold_money_fmt)
        row += 2

        ws.write(row, 0, "", section_fmt)
        ws.write(row, 1, "Other Revenue", section_fmt)
        ws.write(row, 2, "", section_fmt)
        row += 1
        ws.write(row, 0, "3", line_num_fmt)
        ws.write(row, 1, "Investment income", label_fmt)
        ws.write(row, 2, p8["3_investment_income"], money_fmt)
        row += 1
        if p8["8a_fundraising_events_gross"] > 0:
            ws.write(row, 0, "8a", line_num_fmt)
            ws.write(row, 1, "Fundraising events — gross income", label_fmt)
            ws.write(row, 2, p8["8a_fundraising_events_gross"], money_fmt)
            row += 1
            ws.write(row, 0, "8b", line_num_fmt)
            ws.write(row, 1, "Less: direct expenses", label_fmt)
            ws.write(row, 2, p8["8b_fundraising_expenses"], money_fmt)
            row += 1
        ws.write(row, 0, "11", line_num_fmt)
        ws.write(row, 1, "Other revenue", label_fmt)
        ws.write(row, 2, p8["11_other_revenue"], money_fmt)
        row += 2
        ws.write(row, 0, "12", line_num_fmt)
        ws.write(row, 1, "TOTAL REVENUE", total_label_fmt)
        ws.write(row, 2, p8["12_total_revenue"], total_money_fmt)

        # --- Part IX: Functional Expenses ---
        ws = workbook.add_worksheet("Part IX - Expenses")
        ws.set_column(0, 0, 8)
        ws.set_column(1, 1, 50)
        ws.set_column(2, 2, 16)
        ws.set_column(3, 3, 16)
        ws.set_column(4, 4, 16)
        ws.set_column(5, 5, 16)
        row = 0
        ws.write(row, 0, "Form 990 — Part IX: Statement of Functional Expenses", title_fmt)
        row += 2

        ws.write(row, 0, "Line", header_col_fmt)
        ws.write(row, 1, "Description", header_col_fmt)
        ws.write(row, 2, "Total", header_col_fmt)
        ws.write(row, 3, "Program\nServices", header_col_fmt)
        ws.write(row, 4, "Management\n& General", header_col_fmt)
        ws.write(row, 5, "Fundraising", header_col_fmt)
        row += 1

        p9 = data["part9_expenses"]
        expense_line_labels = {
            "5_compensation_current_officers": "Compensation of officers, directors, trustees",
            "7_other_salaries": "Other salaries and wages",
            "11a_management_fees": "Fees for services — management / professional",
            "13_office_expenses": "Office expenses / supplies",
            "14_information_technology": "Information technology",
            "16_occupancy": "Occupancy",
            "17_travel": "Travel",
            "19_other_expenses_a": "Other expenses (program / equipment)",
            "19_other_expenses_b": "Other expenses (fundraising)",
            "19_other_expenses_c": "Other expenses (insurance)",
            "19_other_expenses_d": "Other expenses (miscellaneous)",
            "25_total": "TOTAL FUNCTIONAL EXPENSES",
        }

        for line_key, label in expense_line_labels.items():
            if line_key not in p9:
                continue
            row_data = p9[line_key]
            if row_data["Total"] == 0 and line_key != "25_total":
                continue

            is_total = line_key == "25_total"
            lf = total_label_fmt if is_total else label_fmt
            mf = total_money_fmt if is_total else money_fmt

            line_num = line_key.split("_")[0]
            ws.write(row, 0, line_num, line_num_fmt)
            ws.write(row, 1, label, lf)
            ws.write(row, 2, row_data["Total"], mf)
            ws.write(row, 3, row_data["Program Services"], mf)
            ws.write(row, 4, row_data["Management & General"], mf)
            ws.write(row, 5, row_data["Fundraising"], mf)
            row += 1

        # --- Part X: Balance Sheet ---
        ws = workbook.add_worksheet("Part X - Balance Sheet")
        ws.set_column(0, 0, 8)
        ws.set_column(1, 1, 55)
        ws.set_column(2, 2, 18)
        row = 0
        ws.write(row, 0, "Form 990 — Part X: Balance Sheet", title_fmt)
        row += 2

        p10 = data["part10_balance_sheet"]

        ws.write(row, 0, "", section_fmt)
        ws.write(row, 1, "ASSETS", section_fmt)
        ws.write(row, 2, "", section_fmt)
        row += 1
        asset_lines = [
            ("1", "Cash — non-interest-bearing", "1_cash"),
            ("2", "Savings and temporary cash investments", "2_savings"),
            ("3", "Pledges and grants receivable, net", "3_pledges_receivable"),
            ("4", "Accounts receivable, net", "4_accounts_receivable"),
            ("12", "Investments — other securities", "12_investments_other"),
            ("15", "Other assets", "15_other_assets"),
            ("16", "Total assets", "16_total_assets"),
        ]
        for num, label, key in asset_lines:
            val = p10["assets"].get(key, 0)
            if val != 0 or key == "16_total_assets":
                is_total = key == "16_total_assets"
                ws.write(row, 0, num, line_num_fmt)
                ws.write(row, 1, label, bold_fmt if is_total else label_fmt)
                ws.write(row, 2, val, bold_money_fmt if is_total else money_fmt)
                row += 1

        row += 1
        ws.write(row, 0, "", section_fmt)
        ws.write(row, 1, "LIABILITIES", section_fmt)
        ws.write(row, 2, "", section_fmt)
        row += 1
        liab_lines = [
            ("17", "Accounts payable and accrued expenses", "17_accounts_payable"),
            ("25", "Other liabilities", "25_other_liabilities"),
            ("26", "Total liabilities", "26_total_liabilities"),
        ]
        for num, label, key in liab_lines:
            val = p10["liabilities"].get(key, 0)
            if val != 0 or key == "26_total_liabilities":
                is_total = key == "26_total_liabilities"
                ws.write(row, 0, num, line_num_fmt)
                ws.write(row, 1, label, bold_fmt if is_total else label_fmt)
                ws.write(row, 2, val, bold_money_fmt if is_total else money_fmt)
                row += 1

        row += 1
        ws.write(row, 0, "", section_fmt)
        ws.write(row, 1, "NET ASSETS OR FUND BALANCES", section_fmt)
        ws.write(row, 2, "", section_fmt)
        row += 1
        na_lines = [
            ("27", "Unrestricted net assets", "27_unrestricted"),
            ("28", "Temporarily restricted net assets", "28_temporarily_restricted"),
            ("29", "Permanently restricted net assets", "29_permanently_restricted"),
            ("33", "Total net assets or fund balances", "33_total_net_assets"),
            ("34", "Total liabilities and net assets / fund balances", "34_total_liabilities_and_net_assets"),
        ]
        for num, label, key in na_lines:
            val = p10["net_assets"].get(key, 0)
            if val != 0 or key in ("33_total_net_assets", "34_total_liabilities_and_net_assets"):
                is_total = key in ("33_total_net_assets", "34_total_liabilities_and_net_assets")
                ws.write(row, 0, num, line_num_fmt)
                ws.write(row, 1, label, total_label_fmt if is_total else label_fmt)
                ws.write(row, 2, val, total_money_fmt if is_total else money_fmt)
                row += 1

        # --- Schedule A: Public Support ---
        ws = workbook.add_worksheet("Schedule A - Support")
        ws.set_column(0, 0, 45)
        ws.set_column(1, 1, 18)
        row = 0
        ws.write(row, 0, "Schedule A — Public Charity Status and Public Support", title_fmt)
        row += 2

        sa = data["schedule_a"]
        support_lines = [
            ("Gifts, grants, contributions received", sa["gifts_grants_contributions"]),
            ("Government grants", sa["government_grants"]),
            ("Program service revenue", sa["program_service_revenue"]),
            ("Investment income", sa["investment_income"]),
            ("Fundraising event revenue", sa["fundraising_revenue"]),
            ("Other revenue", sa["other_revenue"]),
        ]
        for label, val in support_lines:
            ws.write(row, 0, label, label_fmt)
            ws.write(row, 1, val, money_fmt)
            row += 1

        row += 1
        ws.write(row, 0, "Total support", bold_fmt)
        ws.write(row, 1, sa["total_support"], bold_money_fmt)
        row += 1
        ws.write(row, 0, "Public support", bold_fmt)
        ws.write(row, 1, sa["public_support"], bold_money_fmt)
        row += 2

        pct_fmt = workbook.add_format({
            "num_format": "0.0%", "bold": True, "font_name": "Calibri", "font_size": 14,
        })
        ws.write(row, 0, "Public support percentage", bold_fmt)
        ws.write(row, 1, sa["public_support_percentage"] / 100, pct_fmt)
        row += 1
        status = "PASSES" if sa["meets_33_percent_test"] else "DOES NOT PASS"
        pass_fmt = workbook.add_format({
            "bold": True, "font_size": 12, "font_name": "Calibri",
            "font_color": "#2E7D32" if sa["meets_33_percent_test"] else "#C62828",
        })
        ws.write(row, 0, f"33-1/3% Public Support Test: {status}", pass_fmt)

    output.seek(0)
    return output.getvalue()

"""
Nonprofit Financial Statement Analyzer
Streamlit application for parsing bank statements and generating
FASB-compliant nonprofit financial statements.
"""

import streamlit as st
import pandas as pd
import plotly.express as px
import plotly.graph_objects as go

from parser import parse_bank_statement, standardize_dataframe, detect_columns
from categorizer import (
    categorize_transaction, get_functional_classification,
    REVENUE_CATEGORIES, EXPENSE_CATEGORIES, ALL_CATEGORIES,
    FUNCTIONAL_CATEGORIES,
)
from statements import generate_all_statements
from exporter import export_to_excel
from report_exporter import export_to_pptx, export_to_pdf
from form990 import generate_form990_data
from archive import (
    save_archive, list_archives, load_archive,
    load_multiple_archives, delete_archive,
)

st.set_page_config(
    page_title="Nonprofit Financial Analyzer",
    page_icon="📊",
    layout="wide",
    initial_sidebar_state="expanded",
)

st.markdown("""
<style>
    .main .block-container { padding-top: 2rem; max-width: 1200px; }
    .metric-card {
        background: linear-gradient(135deg, #667eea 0%, #764ba2 100%);
        padding: 1.2rem; border-radius: 12px; color: white;
        text-align: center; margin-bottom: 0.5rem;
    }
    .metric-card h3 { margin: 0; font-size: 0.85rem; opacity: 0.9; }
    .metric-card h1 { margin: 0.2rem 0 0 0; font-size: 1.6rem; }
    .revenue-card {
        background: linear-gradient(135deg, #11998e 0%, #38ef7d 100%);
    }
    .expense-card {
        background: linear-gradient(135deg, #eb3349 0%, #f45c43 100%);
    }
    .net-card-positive {
        background: linear-gradient(135deg, #667eea 0%, #764ba2 100%);
    }
    .net-card-negative {
        background: linear-gradient(135deg, #e53935 0%, #e35d5b 100%);
    }
    .statement-header {
        border-bottom: 3px double #333; padding-bottom: 0.5rem;
        margin-bottom: 1rem;
    }
    .line-item { display: flex; justify-content: space-between; padding: 0.2rem 0; }
    .line-item-indent { padding-left: 2rem; }
    .line-item-total {
        font-weight: bold; border-top: 1px solid #ccc;
        padding-top: 0.3rem; margin-top: 0.3rem;
    }
    .line-item-grand-total {
        font-weight: bold; border-top: 3px double #333;
        padding-top: 0.3rem; margin-top: 0.3rem;
    }
    .account-badge {
        display: inline-block; padding: 0.15rem 0.6rem; border-radius: 12px;
        font-size: 0.8rem; font-weight: 600; margin-right: 0.3rem;
        background: #e8eaf6; color: #3949ab;
    }
    div[data-testid="stSidebar"] { background-color: #f8f9fa; }
</style>
""", unsafe_allow_html=True)


def format_currency(value: float) -> str:
    """Format a number as currency."""
    if value < 0:
        return f"(${abs(value):,.2f})"
    return f"${value:,.2f}"


def render_metric_card(label: str, value: float, card_class: str = "metric-card"):
    """Render a styled metric card."""
    st.markdown(
        f'<div class="metric-card {card_class}">'
        f'<h3>{label}</h3><h1>{format_currency(value)}</h1></div>',
        unsafe_allow_html=True,
    )


def init_session_state():
    """Initialize session state variables."""
    defaults = {
        "std_df": None,
        "categorized": False,
        "org_name": "My Nonprofit Organization",
        "beginning_cash": 0.0,
        "other_assets": 0.0,
        "liabilities": 0.0,
        "accounts": {},
        "pending_files": [],
        "account_filter": "All Accounts",
        "ein": "",
        "beginning_net_assets": 0.0,
    }
    for key, val in defaults.items():
        if key not in st.session_state:
            st.session_state[key] = val


def get_filtered_df() -> pd.DataFrame:
    """Return the transaction dataframe filtered by the selected account."""
    df = st.session_state.std_df
    if df is None:
        return pd.DataFrame()
    acct = st.session_state.account_filter
    if acct and acct != "All Accounts":
        return df[df["Account"] == acct].copy()
    return df.copy()


def get_account_list() -> list[str]:
    """Return list of all account names in the data."""
    if st.session_state.std_df is not None:
        return sorted(st.session_state.std_df["Account"].unique().tolist())
    return []


def sidebar():
    """Render the sidebar with organization settings and account filter."""
    st.sidebar.title("Nonprofit Financial Analyzer")
    st.sidebar.markdown("---")

    st.sidebar.subheader("Organization Settings")
    st.session_state.org_name = st.sidebar.text_input(
        "Organization Name", value=st.session_state.org_name
    )
    st.session_state.beginning_cash = st.sidebar.number_input(
        "Beginning Cash Balance ($)", value=st.session_state.beginning_cash,
        step=100.0, format="%.2f",
    )
    st.session_state.other_assets = st.sidebar.number_input(
        "Other Assets ($)", value=st.session_state.other_assets,
        step=100.0, format="%.2f",
    )
    st.session_state.liabilities = st.sidebar.number_input(
        "Total Liabilities ($)", value=st.session_state.liabilities,
        step=100.0, format="%.2f",
    )

    with st.sidebar.expander("Form 990 Settings"):
        st.session_state.ein = st.text_input("EIN (XX-XXXXXXX)", value=st.session_state.ein)
        st.session_state.beginning_net_assets = st.number_input(
            "Beginning Net Assets ($)", value=st.session_state.beginning_net_assets,
            step=100.0, format="%.2f",
        )

    if st.session_state.categorized and st.session_state.std_df is not None:
        accounts = get_account_list()
        if len(accounts) > 1:
            st.sidebar.markdown("---")
            st.sidebar.subheader("Account Filter")
            options = ["All Accounts"] + accounts
            current = st.session_state.account_filter
            idx = options.index(current) if current in options else 0
            st.session_state.account_filter = st.sidebar.radio(
                "View data for:", options, index=idx,
            )

            st.sidebar.markdown("---")
            st.sidebar.subheader("Accounts Loaded")
            for acct in accounts:
                acct_df = st.session_state.std_df[st.session_state.std_df["Account"] == acct]
                count = len(acct_df)
                st.sidebar.markdown(
                    f'<span class="account-badge">{acct}</span> {count} transactions',
                    unsafe_allow_html=True,
                )

    st.sidebar.markdown("---")
    st.sidebar.markdown(
        "**How to use:**\n"
        "1. Upload bank statement(s)\n"
        "2. Name each account\n"
        "3. Verify column mapping\n"
        "4. Review categories\n"
        "5. View financial statements\n"
        "6. Export to Excel"
    )


def upload_section():
    """Render the multi-file upload section."""
    st.header("Upload Bank Statements")

    uploaded_files = st.file_uploader(
        "Upload one or more bank statements (CSV or Excel)",
        type=["csv", "xlsx", "xls"],
        accept_multiple_files=True,
        help="Select multiple files to analyze accounts together.",
    )

    if not uploaded_files:
        if st.session_state.categorized:
            st.info("Your data is already loaded. Use the tabs above to view it, or upload new files here.")
        return

    st.markdown("---")
    st.subheader("Configure Each Account")
    st.markdown("Give each file an account name and verify the column mapping.")

    file_configs = []

    for i, uploaded_file in enumerate(uploaded_files):
        with st.expander(f"**File {i + 1}:** {uploaded_file.name}", expanded=True):
            default_name = uploaded_file.name.rsplit(".", 1)[0]
            default_name = default_name.replace("AccountHistory", "Account").strip()

            account_name = st.text_input(
                "Account Name",
                value=default_name,
                key=f"acct_name_{i}",
                help="Give this account a recognizable name (e.g., 'Checking', 'Savings', 'PayPal')",
            )

            with st.spinner(f"Parsing {uploaded_file.name}..."):
                raw_df, mapping = parse_bank_statement(uploaded_file, uploaded_file.name)

            st.success(f"Found {len(raw_df)} rows")

            all_columns = ["(none)"] + list(raw_df.columns)

            col1, col2, col3 = st.columns(3)
            with col1:
                date_col = st.selectbox(
                    "Date Column", all_columns,
                    index=all_columns.index(mapping["date"]) if mapping["date"] else 0,
                    key=f"date_{i}",
                )
            with col2:
                desc_col = st.selectbox(
                    "Description Column", all_columns,
                    index=all_columns.index(mapping["description"]) if mapping["description"] else 0,
                    key=f"desc_{i}",
                )
            with col3:
                amount_col = st.selectbox(
                    "Amount Column (single)", all_columns,
                    index=all_columns.index(mapping["amount"]) if mapping["amount"] else 0,
                    key=f"amount_{i}",
                )

            col4, col5 = st.columns(2)
            with col4:
                debit_col = st.selectbox(
                    "Debit Column (optional)", all_columns,
                    index=all_columns.index(mapping["debit"]) if mapping["debit"] else 0,
                    key=f"debit_{i}",
                )
            with col5:
                credit_col = st.selectbox(
                    "Credit Column (optional)", all_columns,
                    index=all_columns.index(mapping["credit"]) if mapping["credit"] else 0,
                    key=f"credit_{i}",
                )

            updated_mapping = {
                "date": date_col if date_col != "(none)" else None,
                "description": desc_col if desc_col != "(none)" else None,
                "amount": amount_col if amount_col != "(none)" else None,
                "debit": debit_col if debit_col != "(none)" else None,
                "credit": credit_col if credit_col != "(none)" else None,
            }

            with st.expander("Preview Raw Data", expanded=False):
                st.dataframe(raw_df.head(15), use_container_width=True)

            file_configs.append({
                "raw_df": raw_df,
                "mapping": updated_mapping,
                "account_name": account_name,
                "filename": uploaded_file.name,
            })

    st.markdown("---")

    if st.button("Process All Statements", type="primary", use_container_width=True):
        all_dfs = []
        with st.spinner("Processing all accounts..."):
            for config in file_configs:
                std_df = standardize_dataframe(config["raw_df"], config["mapping"])
                std_df["Account"] = config["account_name"]
                std_df["Category"] = std_df.apply(
                    lambda r: categorize_transaction(r["Description"], r["Amount"]), axis=1
                )
                std_df["Functional"] = std_df["Category"].apply(get_functional_classification)
                all_dfs.append(std_df)
                st.session_state.accounts[config["account_name"]] = {
                    "filename": config["filename"],
                    "transaction_count": len(std_df),
                }

            combined = pd.concat(all_dfs, ignore_index=True)
            combined = combined.sort_values("Date").reset_index(drop=True)
            st.session_state.std_df = combined
            st.session_state.categorized = True
            st.session_state.account_filter = "All Accounts"

        total = len(combined)
        acct_count = len(file_configs)
        st.success(f"Processed {total} transactions across {acct_count} account(s)!")
        st.rerun()


def categorization_section():
    """Render the transaction categorization review section."""
    st.header("Transaction Categories")

    df = st.session_state.std_df
    view_df = get_filtered_df()
    accounts = get_account_list()

    acct_label = st.session_state.account_filter
    if acct_label != "All Accounts":
        st.markdown(f'Viewing: <span class="account-badge">{acct_label}</span>', unsafe_allow_html=True)

    st.markdown("Review and adjust auto-assigned categories. Changes update the financial statements.")

    col1, col2, col3 = st.columns([1, 1, 1])
    with col1:
        filter_cat = st.multiselect(
            "Filter by Category",
            ALL_CATEGORIES,
            default=[],
            help="Leave empty to show all",
        )
    with col2:
        search_text = st.text_input("Search Descriptions", "")
    with col3:
        if len(accounts) > 1 and st.session_state.account_filter == "All Accounts":
            filter_acct = st.multiselect(
                "Filter by Account",
                accounts,
                default=[],
                help="Leave empty to show all",
            )
        else:
            filter_acct = []

    filtered = view_df.copy()
    if filter_cat:
        filtered = filtered[filtered["Category"].isin(filter_cat)]
    if search_text:
        filtered = filtered[
            filtered["Description"].str.contains(search_text, case=False, na=False)
        ]
    if filter_acct:
        filtered = filtered[filtered["Account"].isin(filter_acct)]

    st.markdown(f"Showing **{len(filtered)}** of **{len(view_df)}** transactions")

    show_account_col = len(accounts) > 1
    display_cols = ["Date", "Description", "Amount", "Category"]
    if show_account_col:
        display_cols = ["Date", "Account", "Description", "Amount", "Category"]

    col_config = {
        "Date": st.column_config.DateColumn("Date", format="MM/DD/YYYY"),
        "Description": st.column_config.TextColumn("Description", width="large"),
        "Amount": st.column_config.NumberColumn("Amount", format="$%.2f"),
        "Category": st.column_config.SelectboxColumn(
            "Category", options=ALL_CATEGORIES, width="medium", required=True,
        ),
    }
    if show_account_col:
        col_config["Account"] = st.column_config.TextColumn("Account", disabled=True, width="small")

    edited_df = st.data_editor(
        filtered[display_cols].reset_index(drop=True),
        column_config=col_config,
        use_container_width=True,
        num_rows="fixed",
        hide_index=True,
    )

    if st.button("Save Category Changes", type="primary"):
        for idx in range(len(edited_df)):
            orig_idx = filtered.index[idx] if idx < len(filtered) else None
            if orig_idx is not None and orig_idx < len(df):
                new_cat = edited_df.iloc[idx]["Category"]
                st.session_state.std_df.at[orig_idx, "Category"] = new_cat
                st.session_state.std_df.at[orig_idx, "Functional"] = get_functional_classification(new_cat)
        st.success("Categories updated!")
        st.rerun()


def dashboard_section():
    """Render the dashboard with charts and summary metrics."""
    st.header("Financial Dashboard")

    df = get_filtered_df()
    accounts = get_account_list()
    acct_label = st.session_state.account_filter

    if acct_label != "All Accounts":
        st.markdown(f'Viewing: <span class="account-badge">{acct_label}</span>', unsafe_allow_html=True)

    non_transfer = df[df["Category"] != "Internal Account Transfer"]
    revenue = non_transfer[non_transfer["Amount"] > 0]["Amount"].sum()
    expenses = abs(non_transfer[non_transfer["Amount"] < 0]["Amount"].sum())
    net = revenue - expenses
    transaction_count = len(df)

    col1, col2, col3, col4 = st.columns(4)
    with col1:
        render_metric_card("Total Revenue", revenue, "metric-card revenue-card")
    with col2:
        render_metric_card("Total Expenses", expenses, "metric-card expense-card")
    with col3:
        card_class = "metric-card net-card-positive" if net >= 0 else "metric-card net-card-negative"
        render_metric_card("Net Change", net, card_class)
    with col4:
        st.markdown(
            f'<div class="metric-card">'
            f'<h3>Transactions</h3><h1>{transaction_count:,}</h1></div>',
            unsafe_allow_html=True,
        )

    if len(accounts) > 1 and acct_label == "All Accounts":
        st.markdown("---")
        st.subheader("Account Summary")
        acct_rows = []
        full_df = st.session_state.std_df
        for acct in accounts:
            adf = full_df[full_df["Account"] == acct]
            adf_nt = adf[adf["Category"] != "Internal Account Transfer"]
            a_rev = adf_nt[adf_nt["Amount"] > 0]["Amount"].sum()
            a_exp = abs(adf_nt[adf_nt["Amount"] < 0]["Amount"].sum())
            acct_rows.append({
                "Account": acct,
                "Transactions": len(adf),
                "Revenue": round(a_rev, 2),
                "Expenses": round(a_exp, 2),
                "Net": round(a_rev - a_exp, 2),
            })
        acct_summary = pd.DataFrame(acct_rows)
        st.dataframe(
            acct_summary,
            column_config={
                "Account": st.column_config.TextColumn("Account"),
                "Transactions": st.column_config.NumberColumn("Transactions"),
                "Revenue": st.column_config.NumberColumn("Revenue", format="$%.2f"),
                "Expenses": st.column_config.NumberColumn("Expenses", format="$%.2f"),
                "Net": st.column_config.NumberColumn("Net", format="$%.2f"),
            },
            use_container_width=True,
            hide_index=True,
        )

    st.markdown("---")

    col_left, col_right = st.columns(2)

    with col_left:
        st.subheader("Revenue by Category")
        rev_df = non_transfer[non_transfer["Amount"] > 0].groupby("Category")["Amount"].sum().reset_index()
        rev_df.columns = ["Category", "Amount"]
        if len(rev_df) > 0:
            fig_rev = px.pie(
                rev_df, names="Category", values="Amount",
                color_discrete_sequence=px.colors.qualitative.Set3,
                hole=0.4,
            )
            fig_rev.update_traces(textposition="inside", textinfo="percent+label")
            fig_rev.update_layout(
                showlegend=False, margin=dict(t=20, b=20, l=20, r=20), height=350,
            )
            st.plotly_chart(fig_rev, use_container_width=True)
        else:
            st.info("No revenue transactions found.")

    with col_right:
        st.subheader("Expenses by Category")
        exp_df = non_transfer[non_transfer["Amount"] < 0].copy()
        exp_df["AbsAmount"] = exp_df["Amount"].abs()
        exp_by_cat = exp_df.groupby("Category")["AbsAmount"].sum().reset_index()
        exp_by_cat.columns = ["Category", "Amount"]
        if len(exp_by_cat) > 0:
            fig_exp = px.pie(
                exp_by_cat, names="Category", values="Amount",
                color_discrete_sequence=px.colors.qualitative.Pastel,
                hole=0.4,
            )
            fig_exp.update_traces(textposition="inside", textinfo="percent+label")
            fig_exp.update_layout(
                showlegend=False, margin=dict(t=20, b=20, l=20, r=20), height=350,
            )
            st.plotly_chart(fig_exp, use_container_width=True)
        else:
            st.info("No expense transactions found.")

    st.subheader("Monthly Cash Flow")
    monthly = df.copy()
    monthly["Month"] = monthly["Date"].dt.to_period("M").dt.to_timestamp()
    monthly_rev = monthly[monthly["Amount"] > 0].groupby("Month")["Amount"].sum()
    monthly_exp = monthly[monthly["Amount"] < 0].groupby("Month")["Amount"].sum().abs()

    months_index = sorted(set(monthly_rev.index) | set(monthly_exp.index))
    if months_index:
        fig_monthly = go.Figure()
        fig_monthly.add_trace(go.Bar(
            x=[m.strftime("%b %Y") for m in months_index],
            y=[monthly_rev.get(m, 0) for m in months_index],
            name="Revenue",
            marker_color="#38ef7d",
        ))
        fig_monthly.add_trace(go.Bar(
            x=[m.strftime("%b %Y") for m in months_index],
            y=[monthly_exp.get(m, 0) for m in months_index],
            name="Expenses",
            marker_color="#f45c43",
        ))
        fig_monthly.update_layout(
            barmode="group", height=350,
            margin=dict(t=20, b=40, l=40, r=20),
            legend=dict(orientation="h", yanchor="bottom", y=1.02, xanchor="right", x=1),
        )
        st.plotly_chart(fig_monthly, use_container_width=True)

    st.subheader("Functional Expense Allocation")
    func_df = df[df["Amount"] < 0].copy()
    func_df["AbsAmount"] = func_df["Amount"].abs()
    func_df["Functional"] = func_df["Category"].apply(get_functional_classification)
    func_summary = func_df.groupby("Functional")["AbsAmount"].sum().reset_index()
    func_summary.columns = ["Classification", "Amount"]
    func_summary = func_summary[func_summary["Classification"] != "N/A"]

    if len(func_summary) > 0:
        col_a, col_b = st.columns([1, 2])
        with col_a:
            fig_func = px.pie(
                func_summary, names="Classification", values="Amount",
                color="Classification",
                color_discrete_map={
                    "Program Services": "#667eea",
                    "Management & General": "#f093fb",
                    "Fundraising": "#4facfe",
                },
                hole=0.5,
            )
            fig_func.update_traces(textposition="inside", textinfo="percent+label")
            fig_func.update_layout(
                showlegend=False, margin=dict(t=20, b=20, l=20, r=20), height=300,
            )
            st.plotly_chart(fig_func, use_container_width=True)

        with col_b:
            total_exp_func = func_summary["Amount"].sum()
            for _, row in func_summary.iterrows():
                pct = (row["Amount"] / total_exp_func * 100) if total_exp_func > 0 else 0
                st.markdown(
                    f"**{row['Classification']}**: {format_currency(row['Amount'])} ({pct:.1f}%)"
                )
            st.markdown(f"**Total**: {format_currency(total_exp_func)}")

            if total_exp_func > 0:
                program_pct = func_summary.loc[
                    func_summary["Classification"] == "Program Services", "Amount"
                ].sum() / total_exp_func * 100
                if program_pct >= 75:
                    st.success(f"Program expense ratio: {program_pct:.1f}% (Excellent)")
                elif program_pct >= 65:
                    st.warning(f"Program expense ratio: {program_pct:.1f}% (Good)")
                else:
                    st.error(f"Program expense ratio: {program_pct:.1f}% (Below recommended)")


def statements_section():
    """Render the financial statements."""
    st.header("Financial Statements")

    df = get_filtered_df()
    acct_label = st.session_state.account_filter
    accounts = get_account_list()

    if acct_label != "All Accounts":
        st.markdown(f'Generating for: <span class="account-badge">{acct_label}</span>', unsafe_allow_html=True)
    elif len(accounts) > 1:
        st.markdown(f"Generating **combined** statements across **{len(accounts)} accounts**")

    if len(df) == 0:
        st.warning("No transactions to generate statements from.")
        return

    stmts = generate_all_statements(
        df,
        org_name=st.session_state.org_name,
        beginning_cash=st.session_state.beginning_cash,
        other_assets=st.session_state.other_assets,
        liabilities=st.session_state.liabilities,
    )

    excel_bytes = export_to_excel(stmts, df)
    suffix = f"_{acct_label.replace(' ', '_')}" if acct_label != "All Accounts" else ""
    base_name = st.session_state.org_name.replace(" ", "_")

    dl_col1, dl_col2, dl_col3 = st.columns(3)
    with dl_col1:
        st.download_button(
            label="Download Excel",
            data=excel_bytes,
            file_name=f"{base_name}{suffix}_Financial_Statements.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            type="primary",
            use_container_width=True,
        )
    with dl_col2:
        pptx_bytes = export_to_pptx(stmts, df, org_name=st.session_state.org_name)
        st.download_button(
            label="Download PowerPoint",
            data=pptx_bytes,
            file_name=f"{base_name}{suffix}_Financial_Statements.pptx",
            mime="application/vnd.openxmlformats-officedocument.presentationml.presentation",
            type="primary",
            use_container_width=True,
        )
    with dl_col3:
        pdf_bytes = export_to_pdf(stmts, df, org_name=st.session_state.org_name)
        st.download_button(
            label="Download PDF",
            data=pdf_bytes,
            file_name=f"{base_name}{suffix}_Financial_Statements.pdf",
            mime="application/pdf",
            type="primary",
            use_container_width=True,
        )

    st.markdown("---")

    tab1, tab2, tab3, tab4 = st.tabs([
        "Statement of Activities",
        "Statement of Financial Position",
        "Statement of Functional Expenses",
        "Statement of Cash Flows",
    ])

    with tab1:
        _render_activities(stmts["activities"])

    with tab2:
        _render_position(stmts["position"])

    with tab3:
        _render_functional(stmts["functional_expenses"])

    with tab4:
        _render_cash_flows(stmts["cash_flows"])


def _render_activities(data: dict):
    """Render Statement of Activities."""
    st.markdown(f"### {data['title']}")
    st.markdown(f"**{data['organization']}**")
    st.markdown(f"*{data['period']}*")
    st.markdown("---")

    d = data["without_donor_restrictions"]

    st.markdown("**REVENUE AND SUPPORT** *(Without Donor Restrictions)*")
    for label, val in d["revenue"].items():
        col1, col2 = st.columns([3, 1])
        is_sub = ": " in label
        indent = "&emsp;&emsp;&emsp;&emsp;" if is_sub else "&emsp;&emsp;"
        display = label.split(": ", 1)[1] if is_sub else label
        style = "font-size:0.92em; color:#555;" if is_sub else ""
        col1.markdown(f"<span style='{style}'>{indent}{display}</span>", unsafe_allow_html=True)
        col2.markdown(f"<div style='text-align:right;{style}'>{format_currency(val)}</div>", unsafe_allow_html=True)

    col1, col2 = st.columns([3, 1])
    col1.markdown("**Total Revenue and Support**")
    col2.markdown(f"<div style='text-align:right'><b>{format_currency(d['total_revenue'])}</b></div>", unsafe_allow_html=True)

    st.markdown("")
    st.markdown("**EXPENSES**")
    for label, val in d["expenses"].items():
        col1, col2 = st.columns([3, 1])
        is_sub = ": " in label
        indent = "&emsp;&emsp;&emsp;&emsp;" if is_sub else "&emsp;&emsp;"
        display = label.split(": ", 1)[1] if is_sub else label
        style = "font-size:0.92em; color:#555;" if is_sub else ""
        col1.markdown(f"<span style='{style}'>{indent}{display}</span>", unsafe_allow_html=True)
        col2.markdown(f"<div style='text-align:right;{style}'>{format_currency(val)}</div>", unsafe_allow_html=True)

    col1, col2 = st.columns([3, 1])
    col1.markdown("**Total Expenses**")
    col2.markdown(f"<div style='text-align:right'><b>{format_currency(d['total_expenses'])}</b></div>", unsafe_allow_html=True)

    if d.get("net_transfers", 0) != 0:
        st.markdown("")
        st.markdown("**INTERNAL ACCOUNT TRANSFERS (NET)**")
        col1, col2 = st.columns([3, 1])
        col1.markdown("&emsp;&emsp;Net Transfers")
        col2.markdown(f"<div style='text-align:right'>{format_currency(d['net_transfers'])}</div>", unsafe_allow_html=True)

    st.markdown("---")
    col1, col2 = st.columns([3, 1])
    col1.markdown("### Change in Net Assets")
    col2.markdown(f"<div style='text-align:right'><b style='font-size:1.2em'>{format_currency(d['change_in_net_assets'])}</b></div>", unsafe_allow_html=True)


def _render_position(data: dict):
    """Render Statement of Financial Position."""
    st.markdown(f"### {data['title']}")
    st.markdown(f"**{data['organization']}**")
    st.markdown(f"*{data['as_of']}*")
    st.markdown("---")

    st.markdown("**ASSETS**")
    for label, val in data["assets"].items():
        col1, col2 = st.columns([3, 1])
        is_total = "Total" in label
        prefix = "" if is_total else "&emsp;&emsp;"
        weight = "**" if is_total else ""
        col1.markdown(f"{prefix}{weight}{label}{weight}")
        col2.markdown(f"<div style='text-align:right'>{weight}{format_currency(val)}{weight}</div>", unsafe_allow_html=True)

    st.markdown("")
    st.markdown("**LIABILITIES**")
    for label, val in data["liabilities"].items():
        col1, col2 = st.columns([3, 1])
        col1.markdown(f"&emsp;&emsp;{label}")
        col2.markdown(f"<div style='text-align:right'>{format_currency(val)}</div>", unsafe_allow_html=True)

    st.markdown("")
    st.markdown("**NET ASSETS**")
    for label, val in data["net_assets"].items():
        col1, col2 = st.columns([3, 1])
        is_total = "Total" in label
        prefix = "" if is_total else "&emsp;&emsp;"
        weight = "**" if is_total else ""
        col1.markdown(f"{prefix}{weight}{label}{weight}")
        col2.markdown(f"<div style='text-align:right'>{weight}{format_currency(val)}{weight}</div>", unsafe_allow_html=True)

    st.markdown("---")
    col1, col2 = st.columns([3, 1])
    col1.markdown("### Total Liabilities and Net Assets")
    col2.markdown(f"<div style='text-align:right'><b style='font-size:1.2em'>{format_currency(data['total_liabilities_and_net_assets'])}</b></div>", unsafe_allow_html=True)


def _render_functional(data: dict):
    """Render Statement of Functional Expenses."""
    st.markdown(f"### {data['title']}")
    st.markdown(f"**{data['organization']}**")
    st.markdown(f"*{data['period']}*")
    st.markdown("---")

    if not data["table"]:
        st.info("No expense transactions to display.")
        return

    rows = []
    for nat_cat, values in data["table"].items():
        row = {"Expense Category": nat_cat}
        for func_cat in data["functional_categories"]:
            row[func_cat] = values.get(func_cat, 0)
        row["Total"] = values.get("Total", 0)
        rows.append(row)

    totals_row = {"Expense Category": "**Total Expenses**"}
    for func_cat in data["functional_categories"]:
        totals_row[func_cat] = data["totals"].get(func_cat, 0)
    totals_row["Total"] = data["totals"].get("Total", 0)
    rows.append(totals_row)

    table_df = pd.DataFrame(rows)
    money_cols = {c: st.column_config.NumberColumn(c, format="$%.2f")
                  for c in data["functional_categories"] + ["Total"]}

    st.dataframe(
        table_df,
        column_config={
            "Expense Category": st.column_config.TextColumn("Expense Category", width="large"),
            **money_cols,
        },
        use_container_width=True,
        hide_index=True,
    )


def _render_cash_flows(data: dict):
    """Render Statement of Cash Flows."""
    st.markdown(f"### {data['title']}")
    st.markdown(f"**{data['organization']}**")
    st.markdown(f"*{data['period']}*")
    st.markdown("---")

    st.markdown("**CASH FLOWS FROM OPERATING ACTIVITIES**")
    ops = data["operating_activities"]
    for label, val in ops["inflows"].items():
        col1, col2 = st.columns([3, 1])
        col1.markdown(f"&emsp;&emsp;{label}")
        col2.markdown(f"<div style='text-align:right'>{format_currency(val)}</div>", unsafe_allow_html=True)
    for label, val in ops["outflows"].items():
        col1, col2 = st.columns([3, 1])
        col1.markdown(f"&emsp;&emsp;{label}")
        col2.markdown(f"<div style='text-align:right'>{format_currency(val)}</div>", unsafe_allow_html=True)

    col1, col2 = st.columns([3, 1])
    col1.markdown("**Net Cash from Operating Activities**")
    col2.markdown(f"<div style='text-align:right'><b>{format_currency(ops['net'])}</b></div>", unsafe_allow_html=True)

    st.markdown("")
    st.markdown("**CASH FLOWS FROM INVESTING ACTIVITIES**")
    col1, col2 = st.columns([3, 1])
    col1.markdown("**Net Cash from Investing Activities**")
    col2.markdown(f"<div style='text-align:right'><b>{format_currency(data['investing_activities']['net'])}</b></div>", unsafe_allow_html=True)

    st.markdown("")
    st.markdown("**CASH FLOWS FROM FINANCING ACTIVITIES**")
    col1, col2 = st.columns([3, 1])
    col1.markdown("**Net Cash from Financing Activities**")
    col2.markdown(f"<div style='text-align:right'><b>{format_currency(data['financing_activities']['net'])}</b></div>", unsafe_allow_html=True)

    st.markdown("---")

    items = [
        ("Net Change in Cash", data["net_change_in_cash"], True),
        ("Beginning Cash Balance", data["beginning_cash"], False),
    ]
    for label, val, bold in items:
        col1, col2 = st.columns([3, 1])
        w = "**" if bold else ""
        col1.markdown(f"{w}{label}{w}")
        col2.markdown(f"<div style='text-align:right'>{w}{format_currency(val)}{w}</div>", unsafe_allow_html=True)

    col1, col2 = st.columns([3, 1])
    col1.markdown("### Ending Cash Balance")
    col2.markdown(f"<div style='text-align:right'><b style='font-size:1.2em'>{format_currency(data['ending_cash'])}</b></div>", unsafe_allow_html=True)


def form990_section():
    """Render the IRS Form 990 preparation worksheet."""
    st.header("IRS Form 990 Preparation Worksheet")
    st.markdown(
        "This worksheet organizes your financial data into the format required by "
        "IRS Form 990. **This is a preparation tool, not an official filing.** "
        "Bring this to your accountant or use it to complete your actual Form 990."
    )

    df = get_filtered_df()
    if len(df) == 0:
        st.warning("No transactions to generate Form 990 data from.")
        return

    gross_receipts = df[df["Amount"] > 0]["Amount"].sum()

    if gross_receipts <= 50000:
        st.info(
            f"**Gross receipts: {format_currency(gross_receipts)}** — Your gross receipts "
            f"are **$50,000 or less**. You are only required to file **Form 990-N (e-Postcard)**, "
            f"which can be submitted electronically at [IRS.gov](https://www.irs.gov/charities-non-profits/annual-electronic-filing-requirement-for-small-exempt-organizations-form-990-n-e-postcard). "
            f"A full Form 990 is not required, but you may still file one voluntarily."
        )
    elif gross_receipts < 200000:
        st.warning(
            f"**Gross receipts: {format_currency(gross_receipts)}** — Your gross receipts "
            f"are **over $50,000 but under $200,000**. You are required to file **Form 990-EZ** "
            f"or the full **Form 990**. The worksheet below can help prepare either filing."
        )
    else:
        st.error(
            f"**Gross receipts: {format_currency(gross_receipts)}** — Your gross receipts "
            f"are **$200,000 or more**. You are required to file the **full Form 990**. "
            f"Use the worksheet below to organize your data for filing."
        )

    st.markdown("---")

    data = generate_form990_data(
        df,
        org_name=st.session_state.org_name,
        ein=st.session_state.ein,
        beginning_cash=st.session_state.beginning_cash,
        other_assets=st.session_state.other_assets,
        liabilities=st.session_state.liabilities,
        beginning_net_assets=st.session_state.beginning_net_assets,
    )

    from exporter import export_form990_to_excel
    excel_bytes = export_form990_to_excel(data, df)
    st.download_button(
        label="Download Form 990 Worksheet (Excel)",
        data=excel_bytes,
        file_name=f"{st.session_state.org_name.replace(' ', '_')}_Form990_Worksheet.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        type="primary",
        use_container_width=True,
    )

    st.markdown("---")

    header = data["header"]
    col1, col2, col3 = st.columns(3)
    col1.markdown(f"**Organization:** {header['organization_name']}")
    col2.markdown(f"**EIN:** {header['ein'] or 'Not provided'}")
    col3.markdown(f"**Period:** {header['fiscal_year_start']} to {header['fiscal_year_end']}")

    st.markdown("---")

    tab_p1, tab_p8, tab_p9, tab_p10, tab_sa = st.tabs([
        "Part I: Summary",
        "Part VIII: Revenue",
        "Part IX: Expenses",
        "Part X: Balance Sheet",
        "Schedule A: Public Support",
    ])

    with tab_p1:
        _render_990_part1(data["part1_summary"])

    with tab_p8:
        _render_990_part8(data["part8_revenue"])

    with tab_p9:
        _render_990_part9(data["part9_expenses"])

    with tab_p10:
        _render_990_part10(data["part10_balance_sheet"])

    with tab_sa:
        _render_990_schedule_a(data["schedule_a"])


def _990_line(label: str, value: float, line_num: str = "", bold: bool = False):
    """Render a single Form 990 line item."""
    col1, col2, col3 = st.columns([0.5, 3, 1.5])
    w = "**" if bold else ""
    col1.markdown(f"{w}{line_num}{w}")
    col2.markdown(f"{w}{label}{w}")
    col3.markdown(
        f"<div style='text-align:right'>{w}{format_currency(value)}{w}</div>",
        unsafe_allow_html=True,
    )


def _render_990_part1(data: dict):
    """Render Part I: Summary."""
    st.markdown("### Part I — Summary")
    st.markdown("---")

    st.markdown("**Revenue**")
    _990_line("Contributions, gifts, grants", data["line_1"], "1")
    _990_line("Program service revenue", data["line_2"], "2")
    _990_line("Investment income", data["line_3"], "3")
    _990_line("Other revenue", data["line_4"], "4")
    _990_line("**Total revenue**", data["line_12"], "12", bold=True)

    st.markdown("")
    st.markdown("**Expenses**")
    _990_line("Salaries, other compensation, employee benefits", data["line_15"], "15")
    _990_line("Total expenses", data["line_18"], "18", bold=True)

    st.markdown("")
    st.markdown("**Net Assets**")
    _990_line("Revenue less expenses", data["line_19"], "19", bold=True)
    _990_line("Net assets — beginning of year", data["line_20"], "20")
    _990_line("Net assets — end of year", data["line_22"], "22", bold=True)


def _render_990_part8(data: dict):
    """Render Part VIII: Statement of Revenue."""
    st.markdown("### Part VIII — Statement of Revenue")
    st.markdown("---")

    st.markdown("**Contributions, Gifts, Grants (Lines 1a-1h)**")
    _990_line("Government grants (contributions)", data["1e_govt_grants"], "1e")
    _990_line("All other contributions", data["1f_all_other_contributions"], "1f")
    _990_line("Total contributions", data["1h_total_contributions"], "1h", bold=True)

    st.markdown("")
    st.markdown("**Program Service Revenue (Line 2)**")
    if data["2a_program_services"]:
        for desc, val in data["2a_program_services"].items():
            _990_line(desc, val, "2a")
    _990_line("Total program service revenue", data["2a_total"], "2f", bold=True)

    st.markdown("")
    st.markdown("**Other Revenue**")
    _990_line("Investment income", data["3_investment_income"], "3")
    if data["8a_fundraising_events_gross"] > 0:
        _990_line("Fundraising events (gross)", data["8a_fundraising_events_gross"], "8a")
        _990_line("Less: direct expenses", data["8b_fundraising_expenses"], "8b")
        _990_line("Net income from fundraising", data["8c_fundraising_net"], "8c")
    _990_line("Other revenue", data["11_other_revenue"], "11")

    st.markdown("---")
    _990_line("TOTAL REVENUE", data["12_total_revenue"], "12", bold=True)


def _render_990_part9(data: dict):
    """Render Part IX: Statement of Functional Expenses."""
    st.markdown("### Part IX — Statement of Functional Expenses")
    st.markdown("---")

    line_labels = {
        "5_compensation_current_officers": "Compensation of current officers/directors/trustees",
        "7_other_salaries": "Other salaries and wages",
        "11a_management_fees": "Fees for services — management/professional",
        "13_office_expenses": "Office expenses / supplies",
        "14_information_technology": "Information technology",
        "16_occupancy": "Occupancy",
        "17_travel": "Travel",
        "19_other_expenses_a": "Other expenses (program/equipment)",
        "19_other_expenses_b": "Other expenses (fundraising)",
        "19_other_expenses_c": "Other expenses (insurance)",
        "19_other_expenses_d": "Other expenses (miscellaneous)",
        "25_total": "TOTAL FUNCTIONAL EXPENSES",
    }

    header_cols = st.columns([0.4, 2.6, 1, 1, 1, 1])
    headers = ["Line", "Description", "Total", "Program", "Mgmt & General", "Fundraising"]
    for col, h in zip(header_cols, headers):
        col.markdown(f"**{h}**")

    st.markdown("---")

    for line_key, label in line_labels.items():
        if line_key not in data:
            continue
        row = data[line_key]
        if row["Total"] == 0 and line_key != "25_total":
            continue

        is_total = line_key == "25_total"
        w = "**" if is_total else ""
        line_num = line_key.split("_")[0]

        cols = st.columns([0.4, 2.6, 1, 1, 1, 1])
        cols[0].markdown(f"{w}{line_num}{w}")
        cols[1].markdown(f"{w}{label}{w}")
        for i, func in enumerate(["Total", "Program Services", "Management & General", "Fundraising"]):
            cols[i + 2].markdown(
                f"<div style='text-align:right'>{w}{format_currency(row[func])}{w}</div>",
                unsafe_allow_html=True,
            )


def _render_990_part10(data: dict):
    """Render Part X: Balance Sheet."""
    st.markdown("### Part X — Balance Sheet")
    st.markdown("---")

    asset_labels = {
        "1_cash": ("1", "Cash — non-interest-bearing"),
        "2_savings": ("2", "Savings and temporary cash investments"),
        "3_pledges_receivable": ("3", "Pledges and grants receivable"),
        "4_accounts_receivable": ("4", "Accounts receivable"),
        "12_investments_other": ("12", "Investments — other securities"),
        "15_other_assets": ("15", "Other assets"),
        "16_total_assets": ("16", "Total assets"),
    }

    st.markdown("**Assets**")
    for key, (num, label) in asset_labels.items():
        val = data["assets"].get(key, 0)
        if val != 0 or key == "16_total_assets":
            bold = key == "16_total_assets"
            _990_line(label, val, num, bold=bold)

    st.markdown("")
    st.markdown("**Liabilities**")
    liab_labels = {
        "17_accounts_payable": ("17", "Accounts payable and accrued expenses"),
        "25_other_liabilities": ("25", "Other liabilities"),
        "26_total_liabilities": ("26", "Total liabilities"),
    }
    for key, (num, label) in liab_labels.items():
        val = data["liabilities"].get(key, 0)
        if val != 0 or key == "26_total_liabilities":
            bold = key == "26_total_liabilities"
            _990_line(label, val, num, bold=bold)

    st.markdown("")
    st.markdown("**Net Assets / Fund Balances**")
    na_labels = {
        "27_unrestricted": ("27", "Unrestricted net assets"),
        "28_temporarily_restricted": ("28", "Temporarily restricted net assets"),
        "29_permanently_restricted": ("29", "Permanently restricted net assets"),
        "33_total_net_assets": ("33", "Total net assets or fund balances"),
        "34_total_liabilities_and_net_assets": ("34", "Total liabilities and net assets"),
    }
    for key, (num, label) in na_labels.items():
        val = data["net_assets"].get(key, 0)
        if val != 0 or key in ("33_total_net_assets", "34_total_liabilities_and_net_assets"):
            bold = key in ("33_total_net_assets", "34_total_liabilities_and_net_assets")
            _990_line(label, val, num, bold=bold)


def _render_990_schedule_a(data: dict):
    """Render Schedule A: Public Charity Status."""
    st.markdown("### Schedule A — Public Charity Status and Public Support")
    st.markdown("---")

    st.markdown("**Support Schedule**")
    _990_line("Gifts, grants, contributions", data["gifts_grants_contributions"], "")
    _990_line("Government grants", data["government_grants"], "")
    _990_line("Program service revenue", data["program_service_revenue"], "")
    _990_line("Investment income", data["investment_income"], "")
    _990_line("Fundraising event revenue", data["fundraising_revenue"], "")
    _990_line("Other revenue", data["other_revenue"], "")
    _990_line("Total support", data["total_support"], "", bold=True)

    st.markdown("---")
    st.markdown("**Public Support Test (33-1/3%)**")
    _990_line("Public support", data["public_support"], "")

    pct = data["public_support_percentage"]
    col1, col2 = st.columns([3, 1])
    col1.markdown("**Public support percentage**")
    col2.markdown(
        f"<div style='text-align:right'><b>{pct:.1f}%</b></div>",
        unsafe_allow_html=True,
    )

    if data["meets_33_percent_test"]:
        st.success(
            f"Public support is {pct:.1f}% of total support. "
            f"This meets the 33-1/3% public support test for public charity status."
        )
    else:
        st.warning(
            f"Public support is {pct:.1f}% of total support. "
            f"This does NOT meet the 33-1/3% threshold. Review your support sources "
            f"or consult a tax professional about your public charity classification."
        )


def archive_section():
    """Render the archive management section."""
    st.header("Archive Manager")
    st.markdown(
        "Save your current data as a monthly archive, or load previous months "
        "to build cumulative reports (YTD, annual, multi-year)."
    )

    st.markdown("---")

    st.subheader("Save Current Data")
    df = st.session_state.std_df
    if df is not None and len(df) > 0:
        date_min = df["Date"].min().strftime("%b %Y") if pd.notna(df["Date"].min()) else ""
        date_max = df["Date"].max().strftime("%b %Y") if pd.notna(df["Date"].max()) else ""
        default_label = f"{date_min}" if date_min == date_max else f"{date_min} - {date_max}"

        col1, col2 = st.columns([2, 1])
        with col1:
            archive_label = st.text_input(
                "Archive Label",
                value=default_label,
                help="Name this period (e.g., 'January 2026', 'Q1 2026')",
            )
        with col2:
            archive_notes = st.text_input("Notes (optional)", "")

        col_info1, col_info2, col_info3 = st.columns(3)
        col_info1.metric("Transactions", f"{len(df):,}")
        revenue = df[df["Amount"] > 0]["Amount"].sum()
        expenses = abs(df[df["Amount"] < 0]["Amount"].sum())
        col_info2.metric("Revenue", format_currency(revenue))
        col_info3.metric("Expenses", format_currency(expenses))

        if st.button("Save to Archive", type="primary"):
            archive_id = save_archive(
                df,
                label=archive_label,
                org_name=st.session_state.org_name,
                notes=archive_notes,
            )
            st.success(f"Saved archive: **{archive_label}**")
            st.rerun()
    else:
        st.info("No data loaded to archive. Upload a bank statement first.")

    st.markdown("---")

    st.subheader("Saved Archives")
    archives = list_archives()

    if not archives:
        st.info("No archives saved yet. Process a bank statement and save it above.")
        return

    for arch in archives:
        start = arch.get("date_range_start", "")[:10] if arch.get("date_range_start") else "?"
        end = arch.get("date_range_end", "")[:10] if arch.get("date_range_end") else "?"
        created = arch.get("created_at", "")[:16].replace("T", " ")

        with st.expander(f"**{arch['label']}** — {arch['transaction_count']} transactions ({start} to {end})"):
            col1, col2, col3 = st.columns(3)
            col1.markdown(f"**Revenue:** {format_currency(arch.get('total_revenue', 0))}")
            col2.markdown(f"**Expenses:** {format_currency(arch.get('total_expenses', 0))}")
            col3.markdown(f"**Saved:** {created}")

            if arch.get("accounts"):
                st.markdown(f"**Accounts:** {', '.join(arch['accounts'])}")
            if arch.get("notes"):
                st.markdown(f"**Notes:** {arch['notes']}")

            col_a, col_b = st.columns(2)
            with col_a:
                if st.button("Load This Period", key=f"load_{arch['id']}", use_container_width=True):
                    loaded_df, _ = load_archive(arch["id"])
                    st.session_state.std_df = loaded_df
                    st.session_state.categorized = True
                    st.session_state.account_filter = "All Accounts"
                    st.success(f"Loaded **{arch['label']}**")
                    st.rerun()
            with col_b:
                if st.button("Delete", key=f"del_{arch['id']}", type="secondary", use_container_width=True):
                    delete_archive(arch["id"])
                    st.success(f"Deleted **{arch['label']}**")
                    st.rerun()

    st.markdown("---")

    st.subheader("Aggregate Multiple Periods")
    st.markdown("Select multiple archived periods to combine into a cumulative view (e.g., YTD or annual).")

    archive_options = {a["id"]: f"{a['label']} ({a['transaction_count']} txns)" for a in archives}
    selected_ids = st.multiselect(
        "Select periods to combine",
        options=list(archive_options.keys()),
        format_func=lambda x: archive_options[x],
    )

    if selected_ids and st.button("Load & Aggregate Selected Periods", type="primary", use_container_width=True):
        with st.spinner("Loading and combining archives..."):
            combined = load_multiple_archives(selected_ids)

        if len(combined) > 0:
            st.session_state.std_df = combined
            st.session_state.categorized = True
            st.session_state.account_filter = "All Accounts"

            selected_labels = [a["label"] for a in archives if a["id"] in selected_ids]
            st.success(
                f"Loaded **{len(combined)}** transactions from "
                f"**{len(selected_ids)}** periods: {', '.join(selected_labels)}"
            )
            st.rerun()
        else:
            st.warning("No transactions found in the selected archives.")


def _archive_load_section():
    """Render a compact archive loader for the initial upload screen."""
    st.header("Load from Archive")

    archives = list_archives()
    if not archives:
        st.info("No saved archives yet. Upload and process a bank statement to get started.")
        return

    st.markdown("Load a previously saved period:")

    for arch in archives[:5]:
        start = arch.get("date_range_start", "")[:10] if arch.get("date_range_start") else "?"
        end = arch.get("date_range_end", "")[:10] if arch.get("date_range_end") else "?"

        col1, col2 = st.columns([3, 1])
        col1.markdown(
            f"**{arch['label']}** — {arch['transaction_count']} txns "
            f"({start} to {end})"
        )
        with col2:
            if st.button("Load", key=f"init_load_{arch['id']}", use_container_width=True):
                loaded_df, _ = load_archive(arch["id"])
                st.session_state.std_df = loaded_df
                st.session_state.categorized = True
                st.session_state.account_filter = "All Accounts"
                st.rerun()

    if len(archives) > 1:
        st.markdown("---")
        st.markdown("**Or aggregate multiple periods:**")
        archive_options = {a["id"]: a["label"] for a in archives}
        selected = st.multiselect(
            "Select periods",
            options=list(archive_options.keys()),
            format_func=lambda x: archive_options[x],
            key="init_aggregate",
        )
        if selected and st.button("Load & Aggregate", type="primary", key="init_agg_btn"):
            combined = load_multiple_archives(selected)
            if len(combined) > 0:
                st.session_state.std_df = combined
                st.session_state.categorized = True
                st.session_state.account_filter = "All Accounts"
                st.rerun()


def main():
    init_session_state()
    sidebar()

    if st.session_state.std_df is not None and st.session_state.categorized:
        tab_dash, tab_cat, tab_stmt, tab_990, tab_archive, tab_upload = st.tabs([
            "Dashboard", "Categories", "Financial Statements",
            "Form 990 Prep", "Archive", "Upload New",
        ])

        with tab_dash:
            dashboard_section()

        with tab_cat:
            categorization_section()

        with tab_stmt:
            statements_section()

        with tab_990:
            form990_section()

        with tab_archive:
            archive_section()

        with tab_upload:
            upload_section()
    else:
        col_left, col_right = st.columns(2)
        with col_left:
            upload_section()
        with col_right:
            _archive_load_section()


if __name__ == "__main__":
    main()

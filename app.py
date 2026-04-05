import io
import json
import warnings
from pathlib import Path

import pandas as pd
import streamlit as st
from datetime import datetime

from core.categories import (
    DEFAULT_CATEGORY_MAPPING,
    classify_expense,
    load_category_mapping as _load_category_mapping,
    load_budgets,
    save_categories_json,
)
from core.transactions import (
    _parse_excel_dataframe,
    deduplicate_transactions,
)
from core.excel import generate_excel_bytes

# Page setup (must be first Streamlit call)
st.set_page_config(page_title="Personal Budget Analyzer", page_icon="💰", layout="wide")

# Password protection

def check_password():
    try:
        app_password = st.secrets["APP_PASSWORD"]
    except Exception:
        return  # No secrets file — skip password locally
    if "authenticated" not in st.session_state:
        st.session_state.authenticated = False
    if not st.session_state.authenticated:
        st.title("Login")
        password = st.text_input("Password", type="password")
        if st.button("Enter"):
            if password == app_password:
                st.session_state.authenticated = True
                st.rerun()
            else:
                st.error("Wrong password")
        st.stop()

check_password()

warnings.filterwarnings('ignore')


# ==========================================
# App-level wrappers
# ==========================================

@st.cache_data
def load_category_mapping():
    return _load_category_mapping()


def read_excel_file(uploaded_file):
    try:
        preview = pd.read_excel(uploaded_file, header=None, nrows=30)
        header_row = 0
        for i, row in preview.iterrows():
            row_str = row.astype(str).str.cat(sep=' ')
            if "תאריך" in row_str and ("רכישה" in row_str or "עסקה" in row_str):
                header_row = i
                break
        uploaded_file.seek(0)
        df = pd.read_excel(uploaded_file, skiprows=header_row)
        return _parse_excel_dataframe(df, uploaded_file.name)
    except Exception as e:
        st.error(f"Error reading file {uploaded_file.name}: {e}")
        return None


# ==========================================
# User Interface
# ==========================================

st.title("📊 Smart Expense Analysis")
st.markdown("Drag and drop your bank or credit card Excel files, and the system will do the rest.")

# Sidebar
with st.sidebar:
    st.header("Settings")
    mapping = load_category_mapping()
    st.success(f"Loaded {len(mapping)} categories from JSON")
    st.markdown("---")

    # Date Range Filter
    st.subheader("Date Filter")
    date_filter_enabled = st.checkbox("Filter by date range")
    date_from = None
    date_to = None
    if date_filter_enabled:
        date_from = st.date_input("From", value=None)
        date_to = st.date_input("To", value=None)

    st.markdown("---")

    # Categories JSON Editor
    st.subheader("Edit Categories")
    try:
        with open('categories.json', 'r', encoding='utf-8') as f:
            current_json = f.read()
    except FileNotFoundError:
        current_json = json.dumps(
            {**DEFAULT_CATEGORY_MAPPING, "budgets": {"Groceries & Supermarket": 2000}},
            ensure_ascii=False, indent=2
        )

    edited_json = st.text_area("categories.json", value=current_json, height=300)
    if st.button("Save Categories"):
        ok, err = save_categories_json(edited_json)
        if ok:
            st.success("Saved! Reload the page to apply.")
            load_category_mapping.clear()
        else:
            st.error(f"Invalid JSON: {err}")

    st.markdown("---")

    # Categorize a store
    st.subheader("Categorize a Store")
    store_name = st.text_input("Store name (keyword)", placeholder="e.g. ליברה")
    existing_categories = [c for c in mapping.keys() if c != "Uncategorized"]
    cat_options = existing_categories + ["+ New category"]
    selected_cat = st.selectbox("Assign to category", cat_options)
    if selected_cat == "+ New category":
        selected_cat = st.text_input("New category name")
    if st.button("Save Store") and store_name and selected_cat:
        try:
            with open('categories.json', 'r', encoding='utf-8') as f:
                cat_data = json.load(f)
        except FileNotFoundError:
            cat_data = {**DEFAULT_CATEGORY_MAPPING}
        if selected_cat not in cat_data:
            cat_data[selected_cat] = []
        if store_name not in cat_data[selected_cat]:
            cat_data[selected_cat].append(store_name)
        ok, err = save_categories_json(json.dumps(cat_data, ensure_ascii=False, indent=2))
        if ok:
            st.success(f'"{store_name}" added to "{selected_cat}"')
            load_category_mapping.clear()
        else:
            st.error(f"Error: {err}")

    st.markdown("---")
    st.info('Add a `"budgets"` key to set monthly limits per category.\n\nExample:\n```json\n"budgets": {\n  "Groceries": 2000\n}\n```')

class NamedBytesIO(io.BytesIO):
    def __init__(self, data, name):
        super().__init__(data)
        self.name = name


REPORTS_FOLDER = Path(__file__).parent / "Cash_Reports"

# File source selection
source = st.radio("Load files from", ["Folder (Cash_Reports)", "Manual upload"], horizontal=True)

if source == "Folder (Cash_Reports)":
    folder_files = sorted(REPORTS_FOLDER.glob("*.xls*"))
    if not folder_files:
        st.warning(f"No Excel files found in {REPORTS_FOLDER}")
        st.stop()
    st.info(f"Found {len(folder_files)} files in Cash_Reports")
    uploaded_files = [NamedBytesIO(fp.read_bytes(), fp.name) for fp in folder_files]
else:
    uploaded_files = st.file_uploader(
        "Upload Excel files (Multi-select supported)",
        type=['xlsx', 'xls'],
        accept_multiple_files=True,
    )

if uploaded_files:
    all_data = []
    for file in uploaded_files:
        df = read_excel_file(file)
        if df is not None:
            all_data.append(df)

    if all_data:
        full_df = pd.concat(all_data, ignore_index=True)

        # Deduplication
        full_df, n_dupes = deduplicate_transactions(full_df)
        if n_dupes > 0:
            st.warning(f"Removed {n_dupes} duplicate transaction(s) found across uploaded files.")

        # Sort and Classify
        full_df = full_df.sort_values('Date', ascending=False).reset_index(drop=True)
        full_df['Month_Year'] = full_df['Date'].dt.to_period('M')
        full_df['Category'] = full_df['Merchant'].apply(lambda x: classify_expense(x, mapping))

        # Apply date filter
        if date_filter_enabled:
            if date_from:
                full_df = full_df[full_df['Date'] >= pd.Timestamp(date_from)]
            if date_to:
                full_df = full_df[full_df['Date'] <= pd.Timestamp(date_to)]
            if full_df.empty:
                st.warning("No transactions in the selected date range.")
                st.stop()

        # Filters
        col_f1, col_f2 = st.columns(2)
        with col_f1:
            available_months = sorted(full_df['Month_Year'].astype(str).unique(), reverse=True)
            selected_month = st.selectbox("Filter by month", ["All"] + available_months)
        with col_f2:
            store_filter = st.text_input("Filter by store", placeholder="e.g. ליברה")

        if selected_month != "All":
            full_df = full_df[full_df['Month_Year'].astype(str) == selected_month]
        if store_filter:
            full_df = full_df[full_df['Merchant'].str.contains(store_filter, case=False, na=False)]

        # ==========================================
        # TABS
        # ==========================================
        uncategorized_count = (full_df['Category'] == 'Uncategorized').sum()
        tab_overview, tab_deepdive, tab_transactions, tab_categorize = st.tabs(
            ["Overview", "Deep Dive", "Transactions", f"Categorize ({uncategorized_count} left)"]
        )

        # ==========================================
        # TAB 1 — OVERVIEW
        # ==========================================
        with tab_overview:

            # Key Metrics
            total_spent = full_df['Amount'].sum()
            n_months = max(full_df['Month_Year'].nunique(), 1)
            avg_per_month = total_spent / n_months

            col1, col2, col3, col4 = st.columns(4)
            col1.metric("Total Expenses", f"₪{total_spent:,.2f}")
            col2.metric("Avg / Month", f"₪{avg_per_month:,.2f}")
            col3.metric("Total Transactions", len(full_df))
            col4.metric("Active Months", n_months)

            st.divider()

            # Budget Status
            budgets = load_budgets()
            if budgets:
                st.subheader("Budget Status (monthly avg)")
                category_totals = full_df.groupby('Category')['Amount'].sum()
                budget_cols = st.columns(min(len(budgets), 4))
                for i, (cat, limit) in enumerate(budgets.items()):
                    spent = category_totals.get(cat, 0)
                    monthly_avg = spent / n_months
                    pct = min(monthly_avg / limit, 1.0)
                    with budget_cols[i % len(budget_cols)]:
                        st.metric(cat, f"₪{monthly_avg:,.0f} / ₪{limit:,.0f} /mo")
                        st.progress(pct)
                st.divider()

            # Charts: category bar + monthly trend
            col_left, col_right = st.columns(2)

            with col_left:
                st.subheader("Spending by Category")
                category_sum = full_df.groupby('Category')['Amount'].sum().sort_values(ascending=False)
                st.bar_chart(category_sum)

            with col_right:
                st.subheader("Monthly Trend")
                monthly_trend = (
                    full_df.groupby(full_df['Month_Year'].astype(str))['Amount']
                    .sum()
                    .sort_index()
                )
                st.line_chart(monthly_trend)

            st.divider()

            # Top 10 Merchants
            st.subheader("Top 10 Merchants")
            top_merchants = (
                full_df.groupby('Merchant')['Amount']
                .sum()
                .nlargest(10)
                .reset_index()
                .rename(columns={'Amount': 'Total Spent'})
            )
            top_merchants['Total Spent'] = top_merchants['Total Spent'].map(lambda x: f"₪{x:,.2f}")
            st.dataframe(top_merchants, use_container_width=True, hide_index=True)

        # ==========================================
        # TAB 2 — DEEP DIVE
        # ==========================================
        with tab_deepdive:

            # Month-over-Month Change
            st.subheader("Month-over-Month Change")
            mom = (
                full_df.groupby(full_df['Month_Year'].astype(str))['Amount']
                .sum()
                .sort_index()
                .reset_index()
                .rename(columns={'Month_Year': 'Month', 'Amount': 'Total'})
            )
            mom['Prev'] = mom['Total'].shift(1)
            mom['Change (₪)'] = mom['Total'] - mom['Prev']
            mom['Change (%)'] = ((mom['Change (₪)'] / mom['Prev']) * 100).round(1)
            mom['Total'] = mom['Total'].map(lambda x: f"₪{x:,.2f}")
            mom['Change (₪)'] = mom['Change (₪)'].map(lambda x: f"+₪{x:,.2f}" if x > 0 else f"₪{x:,.2f}" if pd.notna(x) else "—")
            mom['Change (%)'] = mom['Change (%)'].map(lambda x: f"+{x}%" if x > 0 else f"{x}%" if pd.notna(x) else "—")
            mom = mom.drop(columns=['Prev'])
            st.dataframe(mom, use_container_width=True, hide_index=True)

            st.divider()

            # Cumulative Spending
            st.subheader("Cumulative Spending Over Time")
            cumulative = (
                full_df.dropna(subset=['Date'])
                .sort_values('Date')
                .assign(Cumulative=lambda d: d['Amount'].cumsum())
                [['Date', 'Cumulative']]
                .set_index('Date')
            )
            st.line_chart(cumulative)

            st.divider()

            # Spending by Day of Week
            st.subheader("Spending by Day of Week")
            dow_order = ['Sunday', 'Monday', 'Tuesday', 'Wednesday', 'Thursday', 'Friday', 'Saturday']
            full_df['DayOfWeek'] = full_df['Date'].dt.day_name()
            dow = (
                full_df.groupby('DayOfWeek')['Amount']
                .sum()
                .reindex(dow_order)
                .dropna()
            )
            st.bar_chart(dow)

            st.divider()

            # Average Transaction Size by Category
            st.subheader("Average Transaction Size by Category")
            avg_tx = (
                full_df.groupby('Category')['Amount']
                .agg(Transactions='count', Average='mean', Total='sum')
                .sort_values('Total', ascending=False)
                .reset_index()
            )
            avg_tx['Average'] = avg_tx['Average'].map(lambda x: f"₪{x:,.2f}")
            avg_tx['Total'] = avg_tx['Total'].map(lambda x: f"₪{x:,.2f}")
            st.dataframe(avg_tx, use_container_width=True, hide_index=True)

            st.divider()

            # Category Drill-Down
            st.subheader("Category Drill-Down")
            all_categories = sorted(full_df['Category'].unique().tolist())
            selected_cat = st.selectbox("Select a category", all_categories)
            cat_df = full_df[full_df['Category'] == selected_cat]

            dc1, dc2, dc3 = st.columns(3)
            dc1.metric("Total Spent", f"₪{cat_df['Amount'].sum():,.2f}")
            dc2.metric("Transactions", len(cat_df))
            dc3.metric("Avg Transaction", f"₪{cat_df['Amount'].mean():,.2f}")

            col_d1, col_d2 = st.columns(2)

            with col_d1:
                st.markdown("**Monthly Breakdown**")
                cat_monthly = (
                    cat_df.groupby(cat_df['Month_Year'].astype(str))['Amount']
                    .sum()
                    .sort_index()
                )
                st.bar_chart(cat_monthly)

            with col_d2:
                st.markdown("**Top Merchants in this Category**")
                cat_merchants = (
                    cat_df.groupby('Merchant')['Amount']
                    .sum()
                    .nlargest(8)
                    .reset_index()
                )
                cat_merchants['Amount'] = cat_merchants['Amount'].map(lambda x: f"₪{x:,.2f}")
                st.dataframe(cat_merchants, use_container_width=True, hide_index=True)

        # ==========================================
        # TAB 3 — TRANSACTIONS
        # ==========================================
        with tab_transactions:

            st.subheader("Review & Re-categorize Transactions")
            filter_cat = st.selectbox(
                "Filter by category",
                ["All"] + sorted(full_df['Category'].unique().tolist()),
                key="tx_filter"
            )
            display_df = full_df if filter_cat == "All" else full_df[full_df['Category'] == filter_cat]

            category_options = list(mapping.keys())
            edited_df = st.data_editor(
                display_df[['Date', 'Merchant', 'Amount', 'Category', 'Source_File']].reset_index(drop=True),
                column_config={
                    "Category": st.column_config.SelectboxColumn(
                        "Category",
                        options=category_options,
                        required=True,
                    )
                },
                use_container_width=True,
                hide_index=True,
                key="category_editor"
            )
            # Sync edits back
            full_df.loc[display_df.index, 'Category'] = edited_df['Category'].values

            # Uncategorized review
            uncategorized = full_df[full_df['Category'] == 'Uncategorized']
            if not uncategorized.empty:
                with st.expander(f"⚠️ {len(uncategorized)} Uncategorized Transactions — click to review"):
                    st.dataframe(
                        uncategorized[['Date', 'Merchant', 'Amount', 'Source_File']],
                        use_container_width=True,
                        hide_index=True
                    )

            st.divider()

            # Pivot Table & Download
            pivot_table = full_df.pivot_table(
                index='Category',
                columns='Month_Year',
                values='Amount',
                aggfunc='sum',
                fill_value=0
            )
            pivot_table['TOTAL'] = pivot_table.sum(axis=1)
            pivot_table.loc['GRAND TOTAL'] = pivot_table.sum()

            excel_data = generate_excel_bytes(full_df, pivot_table)

            st.download_button(
                label="📥 Download Formatted Excel Report",
                data=excel_data,
                file_name=f'Expenses_Report_{datetime.strftime(datetime.today(), "%d_%m_%Y")}.xlsx',
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                type="primary"
            )

        # ==========================================
        # TAB 4 — CATEGORIZE
        # ==========================================
        with tab_categorize:
            uncategorized_merchants = (
                full_df[full_df['Category'] == 'Uncategorized']['Merchant']
                .dropna()
                .unique()
                .tolist()
            )

            if not uncategorized_merchants:
                st.success("All transactions are categorized!")
            else:
                st.subheader(f"{len(uncategorized_merchants)} uncategorized stores")
                st.caption("Assign a category to each store and click Save All.")

                category_options = [c for c in mapping.keys() if c != "Uncategorized"] + ["+ New category"]

                assignments = {}
                for merchant in sorted(uncategorized_merchants):
                    count = (full_df['Merchant'] == merchant).sum()
                    total = full_df[full_df['Merchant'] == merchant]['Amount'].sum()
                    col_m, col_c = st.columns([2, 2])
                    with col_m:
                        st.markdown(f"**{merchant}**  \n{count} transactions · ₪{total:,.0f}")
                    with col_c:
                        choice = st.selectbox("", category_options, key=f"cat_{merchant}", label_visibility="collapsed")
                        if choice == "+ New category":
                            choice = st.text_input("New category name", key=f"new_{merchant}")
                        assignments[merchant] = choice

                if st.button("Save All", type="primary"):
                    try:
                        with open('categories.json', 'r', encoding='utf-8') as f:
                            cat_data = json.load(f)
                    except FileNotFoundError:
                        cat_data = {**DEFAULT_CATEGORY_MAPPING}

                    for merchant, category in assignments.items():
                        if category and category != "+ New category":
                            if category not in cat_data:
                                cat_data[category] = []
                            if merchant not in cat_data[category]:
                                cat_data[category].append(merchant)

                    ok, err = save_categories_json(json.dumps(cat_data, ensure_ascii=False, indent=2))
                    if ok:
                        st.success("Saved! Reload the page to apply.")
                        load_category_mapping.clear()
                    else:
                        st.error(f"Error: {err}")

    else:
        st.warning("No transactions found in the uploaded files.")

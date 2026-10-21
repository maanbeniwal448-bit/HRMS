"""
Comprehensive HR Management System — Streamlit dashboard
Run locally: streamlit run app.py
Dependencies: streamlit, pandas, openpyxl, plotly
"""

from __future__ import annotations

from io import BytesIO
from typing import Optional

import pandas as pd
import plotly.express as px
import streamlit as st

st.set_page_config(
    page_title="Comprehensive HR Management System",
    page_icon="👥",
    layout="wide",
    initial_sidebar_state="expanded",
)

# ---------- Styling ----------
st.markdown(
    """
    <style>
      .block-container {padding-top: 1.5rem; padding-bottom: 2.5rem;}
      [data-testid="stMetric"] {
        background: rgba(128,128,128,0.08);
        border: 1px solid rgba(128,128,128,0.20);
        padding: 16px 18px;
        border-radius: 12px;
      }
      div[data-testid="stDataFrame"] {border-radius: 10px;}
      .small-note {font-size: 0.88rem; opacity: 0.8;}
    </style>
    """,
    unsafe_allow_html=True,
)

st.title("👥 Comprehensive HR Management System")
st.caption("HR Data Analysis Dashboard")
st.markdown("---")


# ---------- Helpers ----------
def normalize_name(value: object) -> str:
    """Normalize a column/sheet name for loose matching."""
    return "".join(ch.lower() for ch in str(value) if ch.isalnum())


def find_column(
    df: pd.DataFrame,
    candidates: list[str],
    contains: Optional[list[str]] = None,
) -> Optional[str]:
    """Find a column by exact normalized name, then by keyword."""
    if df.empty or len(df.columns) == 0:
        return None

    normalized = {normalize_name(col): col for col in df.columns}
    for candidate in candidates:
        key = normalize_name(candidate)
        if key in normalized:
            return normalized[key]

    if contains:
        for col in df.columns:
            key = normalize_name(col)
            if any(normalize_name(word) in key for word in contains):
                return col
    return None


def numeric_columns(df: pd.DataFrame) -> list[str]:
    return df.select_dtypes(include="number").columns.tolist()


def date_columns(df: pd.DataFrame) -> list[str]:
    found = []
    for col in df.columns:
        name = normalize_name(col)
        if any(word in name for word in ("date", "dob", "birth", "joining", "created", "updated")):
            parsed = pd.to_datetime(df[col], errors="coerce", dayfirst=True)
            if parsed.notna().any():
                found.append(col)
    return found


def to_excel_bytes(df: pd.DataFrame) -> bytes:
    output = BytesIO()
    with pd.ExcelWriter(output, engine="openpyxl") as writer:
        df.to_excel(writer, index=False, sheet_name="Filtered Data"[:31])
    return output.getvalue()


def show_metric(label: str, value: object, help_text: Optional[str] = None) -> None:
    st.metric(label=label, value=value, help=help_text)


def make_bar_chart(df: pd.DataFrame, column: str, title: str):
    counts = (
        df[column]
        .fillna("Not specified")
        .astype(str)
        .value_counts()
        .head(20)
        .rename_axis(column)
        .reset_index(name="Records")
    )
    if not counts.empty:
        fig = px.bar(counts, x=column, y="Records", title=title, text="Records")
        fig.update_layout(xaxis_title=column.replace("_", " ").title(), yaxis_title="Records")
        fig.update_traces(textposition="outside")
        st.plotly_chart(fig, use_container_width=True)


def render_sheet_insights(df: pd.DataFrame, sheet_name: str) -> None:
    """Render useful insights without assuming a fixed workbook schema."""
    sheet_key = normalize_name(sheet_name)
    employee_id = find_column(df, ["employee_id", "emp_id", "staff_id"], ["employeeid", "empid"])
    status_col = find_column(df, ["status", "employment_status", "attendance_status"], ["status"])
    department_col = find_column(df, ["department_name", "department", "dept_name"], ["department"])
    salary_col = find_column(df, ["net_salary", "salary", "monthly_salary", "basic_salary", "gross_salary"], ["salary", "wage", "pay"])
    date_col = find_column(df, ["joining_date", "attendance_date", "pay_date", "date", "resignation_date"], ["date"])

    if "employee" in sheet_key:
        c1, c2, c3, c4 = st.columns(4)
        with c1:
            show_metric("Employee Records", f"{len(df):,}")
        with c2:
            if status_col:
                active_mask = df[status_col].astype(str).str.strip().str.lower().isin(
                    ["active", "employed", "working", "permanent"]
                )
                show_metric("Active Employees", f"{int(active_mask.sum()):,}")
            else:
                show_metric("Unique Employee IDs", f"{df[employee_id].nunique(dropna=True):,}" if employee_id else "N/A")
        with c3:
            show_metric("Departments", f"{df[department_col].nunique(dropna=True):,}" if department_col else "N/A")
        with c4:
            show_metric("Duplicate Rows", f"{int(df.duplicated().sum()):,}")

    elif "attendance" in sheet_key:
        c1, c2, c3, c4 = st.columns(4)
        with c1:
            show_metric("Attendance Records", f"{len(df):,}")
        with c2:
            show_metric("Employees Tracked", f"{df[employee_id].nunique(dropna=True):,}" if employee_id else "N/A")
        with c3:
            if status_col:
                present = df[status_col].astype(str).str.lower().str.contains("present|on time", regex=True, na=False).sum()
                show_metric("Present Records", f"{int(present):,}")
            else:
                show_metric("Unique Statuses", f"{df[status_col].nunique(dropna=True):,}" if status_col else "N/A")
        with c4:
            show_metric("Duplicate Rows", f"{int(df.duplicated().sum()):,}")

    elif any(word in sheet_key for word in ("salary", "payroll")):
        c1, c2, c3, c4 = st.columns(4)
        nums = numeric_columns(df)
        with c1:
            show_metric("Payroll Records", f"{len(df):,}")
        with c2:
            show_metric("Employees", f"{df[employee_id].nunique(dropna=True):,}" if employee_id else "N/A")
        with c3:
            if salary_col and pd.to_numeric(df[salary_col], errors="coerce").notna().any():
                total = pd.to_numeric(df[salary_col], errors="coerce").sum()
                show_metric("Total " + salary_col.replace("_", " ").title(), f"{total:,.2f}")
            else:
                show_metric("Numeric Fields", f"{len(nums):,}")
        with c4:
            show_metric("Duplicate Rows", f"{int(df.duplicated().sum()):,}")

    elif "department" in sheet_key:
        c1, c2, c3 = st.columns(3)
        with c1:
            show_metric("Departments", f"{len(df):,}")
        with c2:
            show_metric("Columns", f"{len(df.columns):,}")
        with c3:
            show_metric("Duplicate Rows", f"{int(df.duplicated().sum()):,}")

    elif any(word in sheet_key for word in ("resignation", "exit", "termination")):
        c1, c2, c3, c4 = st.columns(4)
        with c1:
            show_metric("Exit Records", f"{len(df):,}")
        with c2:
            show_metric("Employees", f"{df[employee_id].nunique(dropna=True):,}" if employee_id else "N/A")
        with c3:
            show_metric("Statuses", f"{df[status_col].nunique(dropna=True):,}" if status_col else "N/A")
        with c4:
            show_metric("Duplicate Rows", f"{int(df.duplicated().sum()):,}")

    else:
        c1, c2, c3 = st.columns(3)
        with c1:
            show_metric("Total Records", f"{len(df):,}")
        with c2:
            show_metric("Total Columns", f"{len(df.columns):,}")
        with c3:
            show_metric("Duplicate Rows", f"{int(df.duplicated().sum()):,}")

    # Generic visual insights
    st.subheader("Data Insights")
    chart_cols = st.columns(2)
    categorical = [
        col for col in df.columns
        if (df[col].dtype == "object" or str(df[col].dtype) == "category")
        and 1 < df[col].nunique(dropna=True) <= 50
    ]
    with chart_cols[0]:
        preferred = department_col or status_col
        if preferred:
            make_bar_chart(df, preferred, f"Records by {preferred.replace('_', ' ').title()}")
        elif categorical:
            make_bar_chart(df, categorical[0], f"Records by {categorical[0].replace('_', ' ').title()}")
        else:
            st.info("No suitable categorical column was found for a category chart.")

    with chart_cols[1]:
        if salary_col and pd.to_numeric(df[salary_col], errors="coerce").notna().any():
            salary_data = df[[salary_col]].copy()
            salary_data[salary_col] = pd.to_numeric(salary_data[salary_col], errors="coerce")
            salary_data = salary_data.dropna()
            fig = px.histogram(salary_data, x=salary_col, nbins=20, title=f"Distribution of {salary_col.replace('_', ' ').title()}")
            st.plotly_chart(fig, use_container_width=True)
        elif date_col:
            dates = pd.to_datetime(df[date_col], errors="coerce", dayfirst=True).dropna()
            if not dates.empty:
                trend = dates.dt.to_period("M").astype(str).value_counts().sort_index().rename_axis("Month").reset_index(name="Records")
                fig = px.line(trend, x="Month", y="Records", markers=True, title=f"Records over time ({date_col})")
                st.plotly_chart(fig, use_container_width=True)
            else:
                st.info("No valid dates were found for a time-trend chart.")
        elif len(numeric_columns(df)) > 0:
            num_col = numeric_columns(df)[0]
            fig = px.histogram(df, x=num_col, nbins=20, title=f"Distribution of {num_col.replace('_', ' ').title()}")
            st.plotly_chart(fig, use_container_width=True)
        elif len(categorical) > 1:
            make_bar_chart(df, categorical[1], f"Records by {categorical[1].replace('_', ' ').title()}")
        else:
            st.info("No additional numeric, date, or category field is available for a second chart.")


# ---------- Upload and workbook handling ----------
st.subheader("Upload your HRMS Excel dataset")
uploaded_file = st.file_uploader(
    "Choose an Excel workbook (.xlsx or .xls)",
    type=["xlsx", "xls"],
    help="Upload the workbook containing sheets such as Employees, Departments, Attendance, Salaries, or Resignations.",
)

if uploaded_file is None:
    st.info("Upload your HRMS Excel workbook to view the dashboard.")
    st.markdown(
        """
        **This dashboard supports multiple Excel sheets automatically.**

        It calculates record counts, column counts, duplicate rows, category breakdowns, and available salary/date trends based on the selected sheet.
        """
    )
    st.stop()

try:
    # Read workbook once, preserving sheet names and dataframes.
    workbook = pd.ExcelFile(uploaded_file)
    sheet_names = workbook.sheet_names
    if not sheet_names:
        st.error("The uploaded Excel workbook does not contain any sheets.")
        st.stop()

    # IMPORTANT: define this variable directly from the selectbox to prevent
    # NameError errors involving selected_sheet / selected_sheet_name.
    selected_sheet_name = st.selectbox("Select Excel sheet", options=sheet_names, index=0)

    try:
        df = pd.read_excel(workbook, sheet_name=selected_sheet_name)
    except Exception as exc:
        st.error(f"Could not read sheet '{selected_sheet_name}': {exc}")
        st.stop()

except Exception as exc:
    st.error(f"Unable to open this Excel workbook. Please verify that it is a valid Excel file. Details: {exc}")
    st.stop()

if df is None or df.empty and len(df.columns) == 0:
    st.warning(f"The sheet '{selected_sheet_name}' is empty.")
    st.stop()

# Clean column labels and make duplicate/blank headers usable.
df = df.copy()
df.columns = [
    str(col).strip() if str(col).strip() and not str(col).lower().startswith("unnamed:") else f"Column_{i + 1}"
    for i, col in enumerate(df.columns)
]
df = df.dropna(how="all").reset_index(drop=True)

st.markdown("---")
st.header("HR Overview")
st.subheader(f"HR Analytics Dashboard — {selected_sheet_name}")

# Generic top-level metrics, always shown for every sheet.
total_records = len(df)
total_columns = len(df.columns)
duplicate_rows = int(df.duplicated().sum())
m1, m2, m3 = st.columns(3)
with m1:
    show_metric("Total Records", f"{total_records:,}")
with m2:
    show_metric("Total Columns", f"{total_columns:,}")
with m3:
    show_metric("Duplicate Rows", f"{duplicate_rows:,}")

# Sheet-aware KPIs and charts.
render_sheet_insights(df, selected_sheet_name)

# ---------- Data preview, filters and export ----------
st.markdown("---")
st.subheader("Dataset Preview")

filtered_df = df.copy()
with st.expander("Filter the displayed data", expanded=False):
    filter_columns = st.multiselect(
        "Choose columns to filter",
        options=df.columns.tolist(),
        default=[],
        key=f"filter_columns_{selected_sheet_name}",
    )
    filtered_df = df.copy()
    for col in filter_columns:
        non_null_values = filtered_df[col].dropna().unique().tolist()
        if len(non_null_values) > 200:
            st.caption(f"Skipping '{col}' because it contains more than 200 distinct values.")
            continue
        choices = sorted([str(value) for value in non_null_values])
        chosen = st.multiselect(
            f"Values for {col}",
            options=choices,
            default=choices,
            key=f"filter_values_{selected_sheet_name}_{col}",
        )
        if chosen:
            filtered_df = filtered_df[filtered_df[col].astype(str).isin(chosen)]
        else:
            filtered_df = filtered_df.iloc[0:0]

st.caption(f"Showing {len(filtered_df):,} of {len(df):,} rows.")
st.dataframe(filtered_df, use_container_width=True, hide_index=True)

download_col1, download_col2 = st.columns(2)
with download_col1:
    st.download_button(
        label="⬇️ Download current sheet as CSV",
        data=filtered_df.to_csv(index=False).encode("utf-8-sig"),
        file_name=f"{selected_sheet_name.replace(' ', '_')}_filtered.csv",
        mime="text/csv",
        use_container_width=True,
    )
with download_col2:
    st.download_button(
        label="⬇️ Download current sheet as Excel",
        data=to_excel_bytes(filtered_df),
        file_name=f"{selected_sheet_name.replace(' ', '_')}_filtered.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        use_container_width=True,
    )

# ---------- Data quality ----------
with st.expander("Data quality checks", expanded=False):
    missing = df.isna().sum()
    quality = pd.DataFrame({
        "Column": df.columns,
        "Missing Values": [int(missing[col]) for col in df.columns],
        "Missing (%)": [round(float(missing[col]) / max(len(df), 1) * 100, 2) for col in df.columns],
        "Unique Values": [int(df[col].nunique(dropna=True)) for col in df.columns],
        "Data Type": [str(df[col].dtype) for col in df.columns],
    })
    st.dataframe(quality, use_container_width=True, hide_index=True)
    missing_total = int(df.isna().sum().sum())
    q1, q2, q3 = st.columns(3)
    with q1:
        show_metric("Missing Cells", f"{missing_total:,}")
    with q2:
        show_metric("Duplicate Rows", f"{int(df.duplicated().sum()):,}")
    with q3:
        completeness = 100 * (1 - missing_total / max(df.shape[0] * max(df.shape[1], 1), 1))
        show_metric("Cell Completeness", f"{completeness:.1f}%")

st.markdown("---")
st.caption("Comprehensive HR Management System • Streamlit dashboard")

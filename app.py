import streamlit as st
import pandas as pd
import matplotlib.pyplot as plt

st.set_page_config(
    page_title="HRMS Dashboard",
    page_icon="👥",
    layout="wide"
)

st.title("Comprehensive HR Management System")
st.write("HR Data Analysis Dashboard")

uploaded_file = st.file_uploader(
    "Upload your HRMS Excel dataset",
    type=["xlsx"]
)

if uploaded_file is not None:
    
    # ==========================================
    # 1. INTERACTIVE EMPLOYEE FILTERS
    # ==========================================
   


st.title("Comprehensive HR Management System")

uploaded_file = st.file_uploader(
    "Upload your HR Excel file",
    type=["xlsx", "xls", "csv"]
)

if uploaded_file is not None:
    if uploaded_file.name.endswith(".csv"):
        df = pd.read_csv(uploaded_file)
    else:
        df = pd.read_excel(uploaded_file)

    st.subheader("Employee Data Preview")
    st.dataframe(df.head(), width="stretch")

    # Put your filter and download code here,
    # indented by four spaces inside this block.

    # ==========================================
    # 2. DOWNLOAD FILTERED DATA AS CSV
    # ==========================================
    st.header("Download HR Report")

    csv_data = filtered_df.to_csv(index=False).encode("utf-8")

    st.download_button(
        label="Download Filtered Employee Data (CSV)",
        data=csv_data,
        file_name="HRMS_Filtered_Report.csv",
        mime="text/csv"
    )

    excel_data = pd.ExcelFile(uploaded_file)

    sheet_name = st.selectbox(
        "Select Excel sheet",
        excel_data.sheet_names
    )

    df = pd.read_excel(
        uploaded_file,
        sheet_name=sheet_name
    )
    
    # HR employee metrics
    st.subheader("HR Overview")

   
    st.subheader("HR Analytics Dashboard")

    # Basic employee and data metrics
    col1, col2, col3 = st.columns(3)

    col1.metric("Total Records", len(df))
    col2.metric("Total Columns", len(df.columns))
    col3.metric(
        "Duplicate Rows",
        int(df.duplicated().sum())
    )

    st.subheader("Dataset Preview")
    st.dataframe(df.head(10), use_container_width=True)

    # Find relevant columns automatically
    department_col = next(
        (c for c in df.columns
         if "department" in c.lower()),
        None
    )

    salary_col = next(
        (c for c in df.columns
         if "salary" in c.lower() and
         pd.api.types.is_numeric_dtype(df[c])),
        None
    )

    if department_col:
        st.subheader("Department-wise Employee Count")

        department_counts = (
            df[department_col]
            .fillna("Unknown")
            .value_counts()
        )

        st.bar_chart(department_counts)

    if salary_col:
        st.subheader("Salary Analysis")

        avg_salary = df[salary_col].mean()
        min_salary = df[salary_col].min()
        max_salary = df[salary_col].max()

        col1, col2, col3 = st.columns(3)
        col1.metric("Average Salary", f"{avg_salary:,.2f}")
        col2.metric("Minimum Salary", f"{min_salary:,.2f}")
        col3.metric("Maximum Salary", f"{max_salary:,.2f}")

        st.bar_chart(df[salary_col].dropna())

    st.subheader("Missing Values Analysis")

    missing_values = (
        df.isnull().sum()
        .sort_values(ascending=False)
    )

    st.dataframe(
        missing_values.rename("Missing Values"),
        use_container_width=True
    )

    st.subheader("Numerical Data Analysis")

    numeric_columns = df.select_dtypes(
        include="number"
    ).columns.tolist()

    if numeric_columns:
        selected_column = st.selectbox(
            "Select a numerical column",
            numeric_columns
        )

        st.write(df[selected_column].describe())

        st.bar_chart(
            df[selected_column].dropna()
            .reset_index(drop=True)
        )

   
           




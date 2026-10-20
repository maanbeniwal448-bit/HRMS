
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

    df.columns = df.columns.str.strip()

    def find_column(possible_names):
        column_map = {
            col.lower().replace(" ", "_"): col
            for col in df.columns
        }

        for name in possible_names:
            key = name.lower().replace(" ", "_")
            if key in column_map:
                return column_map[key]

        return None

    employee_col = find_column([
        "employee_id", "emp_id", "employeeid"
    ])
    salary_col = find_column([
        "salary", "monthly_salary", "annual_salary"
    ])
    age_col = find_column([
        "age", "employee_age"
    ])
    department_col = find_column([
        "department", "department_name"
    ])

    total_employees = (
        df[employee_col].nunique()
        if employee_col else len(df)
    )

    c1, c2, c3, c4 = st.columns(4)

    c1.metric("Total Employees", total_employees)

    if salary_col:
        avg_salary = pd.to_numeric(
            df[salary_col], errors="coerce"
        ).mean()
        c2.metric(
            "Average Salary",
            f"{avg_salary:,.2f}"
            if pd.notna(avg_salary) else "N/A"
        )
    else:
        c2.metric("Average Salary", "Column not found")

    if age_col:
        avg_age = pd.to_numeric(
            df[age_col], errors="coerce"
        ).mean()
        c3.metric(
            "Average Age",
            f"{avg_age:.1f} years"
            if pd.notna(avg_age) else "N/A"
        )
    else:
        c3.metric("Average Age", "Column not found")

    if department_col:
        c4.metric(
            "Departments",
            df[department_col].nunique()
        )
    else:
        c4.metric("Departments", "Column not found")


    st.subheader("Dataset Preview")
    st.dataframe(df.head(10), use_container_width=True)

    col1, col2, col3 = st.columns(3)

    col1.metric("Total Rows", len(df))
    col2.metric("Total Columns", len(df.columns))
    col3.metric(
        "Duplicate Rows",
        int(df.duplicated().sum())
    )

    st.subheader("Missing Values")
    st.dataframe(
        df.isnull().sum().rename("Missing Values")
    )

    numeric_columns = df.select_dtypes(
        include="number"
    ).columns.tolist()

    if numeric_columns:
        selected_column = st.selectbox(
            "Select a column for analysis",
            numeric_columns
        )

        fig, ax = plt.subplots()
        df[selected_column].dropna().plot(
            kind="hist", bins=20, ax=ax
        )
        ax.set_title(f"Distribution of {selected_column}")
        st.pyplot(fig)
        plt.close(fig)

else:
    st.info("Upload your Excel dataset to start.")
    
   

   
           


"""
Comprehensive HR Management System (HRMS)
Run: streamlit run app.py

Uses an uploaded Excel workbook when provided. Otherwise, choose the clearly-labelled
synthetic demo dataset to explore the dashboard. Demo records are fictional.
"""
from __future__ import annotations

from io import BytesIO
from datetime import date, timedelta
import re

import pandas as pd
import plotly.express as px
import streamlit as st

st.set_page_config(
    page_title="PeopleDesk | HRMS",
    page_icon="👥",
    layout="wide",
    initial_sidebar_state="expanded",
)

# ---------------------------- Styling ---------------------------------------
st.markdown("""
<style>
.block-container {padding-top: 1.2rem; padding-bottom: 2.2rem; max-width: 1500px;}
[data-testid="stMetric"] {background: rgba(128,128,128,.07); border: 1px solid rgba(128,128,128,.18);
  padding: 15px 17px; border-radius: 13px;}
[data-testid="stMetricLabel"] {font-size: .88rem;}
[data-testid="stMetricValue"] {font-weight: 700;}
div[data-testid="stDataFrame"] {border: 1px solid rgba(128,128,128,.15); border-radius: 10px;}
.stDownloadButton button {border-radius: 9px; font-weight: 600;}
section[data-testid="stSidebar"] {border-right: 1px solid rgba(128,128,128,.15);}
[data-testid="stPlotlyChart"] {border: 1px solid rgba(128,128,128,.12); border-radius: 12px; padding: 6px;}
.hero {padding: 1.1rem 1.3rem; border: 1px solid rgba(128,128,128,.2); border-radius: 16px;
  background: linear-gradient(120deg, rgba(70,110,180,.13), rgba(40,170,150,.07)); margin-bottom: 1rem;}
.hero h1 {margin: 0; font-size: 1.8rem;}
.muted {opacity: .72; font-size: .9rem;}
</style>
""", unsafe_allow_html=True)


# ---------------------------- Helpers ---------------------------------------
def norm(value) -> str:
    return re.sub(r"[^a-z0-9]", "", str(value).lower())


def find_col(frame: pd.DataFrame, exact=(), contains=()):
    if frame is None or frame.empty:
        cols = list(frame.columns) if frame is not None else []
    else:
        cols = list(frame.columns)
    by_norm = {norm(c): c for c in cols}
    for name in exact:
        if norm(name) in by_norm:
            return by_norm[norm(name)]
    for col in cols:
        key = norm(col)
        if any(norm(word) in key for word in contains):
            return col
    return None


def money(value):
    try:
        return f"₹{float(value):,.0f}"
    except (TypeError, ValueError):
        return "—"


@st.cache_data(show_spinner=False, max_entries=8)
def excel_bytes(sheets: dict[str, pd.DataFrame]) -> bytes:
    buff = BytesIO()
    with pd.ExcelWriter(buff, engine="openpyxl") as writer:
        for name, frame in sheets.items():
            frame.to_excel(writer, sheet_name=str(name)[:31] or "Data", index=False)
    return buff.getvalue()


@st.cache_data(show_spinner=False, max_entries=16)
def csv_bytes(frame: pd.DataFrame) -> bytes:
    return frame.to_csv(index=False).encode("utf-8-sig")


def safe_date_series(frame, col):
    if not col:
        return pd.Series(pd.NaT, index=frame.index)
    return pd.to_datetime(frame[col], errors="coerce", dayfirst=True)


def metric_row(items):
    cols = st.columns(len(items))
    for col, (label, value, help_text) in zip(cols, items):
        with col:
            st.metric(label, value, help=help_text)


def clean_frame(frame):
    frame = frame.copy()
    frame.columns = [
        str(c).strip() if str(c).strip() and not str(c).lower().startswith("unnamed:")
        else f"Column_{i+1}"
        for i, c in enumerate(frame.columns)
    ]
    # Make duplicate column names unique without discarding data.
    seen, new_cols = {}, []
    for c in frame.columns:
        seen[c] = seen.get(c, 0) + 1
        new_cols.append(c if seen[c] == 1 else f"{c}_{seen[c]}")
    frame.columns = new_cols
    return frame.dropna(how="all").reset_index(drop=True)


@st.cache_data(show_spinner=False, max_entries=1)
def demo_data():
    """Reproducible fictional sample data; never presented as real company data."""
    import numpy as np
    rng = np.random.default_rng(42)
    first = ["Aarav", "Aditi", "Vihaan", "Ananya", "Arjun", "Ishita", "Reyansh", "Meera",
             "Kabir", "Saanvi", "Advik", "Diya", "Rohan", "Myra", "Krishna", "Sara",
             "Ishaan", "Kiara", "Dev", "Tara", "Neil", "Anika", "Yash", "Riya"]
    last = ["Sharma", "Verma", "Singh", "Patel", "Gupta", "Kumar", "Mehta", "Reddy",
            "Joshi", "Malhotra", "Nair", "Das"]
    departments = ["Engineering", "Human Resources", "Finance", "Sales", "Marketing",
                   "Operations", "Customer Success", "Product"]
    titles = {
        "Engineering": ["Software Engineer", "Senior Developer", "QA Engineer", "Tech Lead"],
        "Human Resources": ["HR Executive", "HR Business Partner", "Talent Specialist"],
        "Finance": ["Accountant", "Financial Analyst", "Finance Manager"],
        "Sales": ["Sales Executive", "Account Manager", "Sales Manager"],
        "Marketing": ["Marketing Executive", "Content Strategist", "Growth Manager"],
        "Operations": ["Operations Associate", "Operations Manager", "Process Analyst"],
        "Customer Success": ["Support Specialist", "Customer Success Manager"],
        "Product": ["Product Analyst", "Product Manager", "UX Designer"],
    }
    n = 180
    names = [f"{first[i % len(first)]} {last[(i * 7 + i // 5) % len(last)]}" for i in range(n)]
    dept = rng.choice(departments, n, p=[.28,.08,.10,.16,.12,.12,.07,.07])
    statuses = rng.choice(["Active", "Active", "Active", "Active", "On Leave", "Inactive"], n,
                          p=[.64,.12,.08,.06,.06,.04])
    today = pd.Timestamp.today().normalize()
    join_start = today - pd.Timedelta(days=7 * 365)
    join_days = rng.integers(0, 7 * 365 + 1, n)
    employees = pd.DataFrame({
        "employee_id": [f"EMP{i+1001}" for i in range(n)],
        "employee_name": names,
        "email": [f"{norm(x)}{i+1001}@peopledesk.example" for i, x in enumerate(names)],
        "department": dept,
        "job_title": [str(rng.choice(titles[d])) for d in dept],
        "location": rng.choice(["Delhi", "Bengaluru", "Mumbai", "Hyderabad", "Pune"], n),
        "joining_date": (join_start + pd.to_timedelta(join_days, unit="D")).date,
        "employment_type": rng.choice(["Full-time", "Full-time", "Full-time", "Contract", "Intern"], n),
        "status": statuses,
        "manager": [f"{first[(i+3)%len(first)]} {last[(i+2)%len(last)]}" for i in range(n)],
        "annual_salary": rng.integers(360000, 2400000, n).round(-3),
    })
    employees["age"] = rng.integers(22, 59, n)
    employees["joining_date"] = pd.to_datetime(employees["joining_date"])
    active_ids = employees.loc[employees.status.eq("Active"), "employee_id"].tolist()
    att_rows = []
    today = pd.Timestamp.today().normalize()
    for day_offset in range(29, -1, -1):
        day = today - pd.Timedelta(days=day_offset)
        if day.weekday() >= 5:
            continue
        for emp in active_ids:
            status = rng.choice(["Present", "Present", "Present", "Work From Home", "Absent", "Half Day"],
                                p=[.68,.10,.06,.08,.04,.04])
            att_rows.append((emp, day, status, round(float(rng.normal(8.2, .7)), 1) if status not in ("Absent",) else 0))
    attendance = pd.DataFrame(att_rows, columns=["employee_id", "attendance_date", "attendance_status", "hours_worked"])
    # Payroll sample includes current employees only and reconciles gross/net components.
    payroll_employees = employees[employees["status"].isin(["Active", "On Leave"])][
        ["employee_id", "employee_name", "department", "annual_salary"]
    ].copy()
    salaries = payroll_employees
    salaries["pay_month"] = today.replace(day=1)
    salaries["gross_salary"] = (salaries["annual_salary"] / 12).round(0)
    salaries["basic_salary"] = (salaries["gross_salary"] * .50).round(0)
    salaries["allowances"] = (salaries["gross_salary"] * .30).round(0)
    salaries["deductions"] = (salaries["gross_salary"] * .08).round(0)
    salaries["net_salary"] = (salaries["basic_salary"] + salaries["allowances"] - salaries["deductions"]).round(0)
    salaries["payroll_status"] = rng.choice(["Processed", "Pending", "On Hold"], len(salaries), p=[.94, .04, .02])
    leaves = []
    leave_types = ["Annual Leave", "Sick Leave", "Personal Leave", "Compensatory Off"]
    for i in range(45):
        emp = str(rng.choice(employees.employee_id))
        start = today + pd.Timedelta(days=int(rng.integers(-15, 28)))
        leaves.append({
            "request_id": f"LV-{3001+i}", "employee_id": emp,
            "employee_name": employees.loc[employees.employee_id.eq(emp), "employee_name"].iloc[0],
            "leave_type": str(rng.choice(leave_types)),
            "start_date": start, "end_date": start + pd.Timedelta(days=int(rng.integers(0, 4))),
            "days": int(rng.integers(1, 5)),
            "status": str(rng.choice(["Pending", "Approved", "Approved", "Rejected"], p=[.28,.52,.12,.08])),
        })
    recruitment = pd.DataFrame({
        "requisition_id": [f"REQ-{410+i}" for i in range(12)],
        "job_title": ["Software Engineer", "HR Executive", "Data Analyst", "Sales Manager",
                      "Product Designer", "Accountant", "QA Engineer", "Marketing Executive",
                      "DevOps Engineer", "Recruiter", "Customer Success Associate", "Business Analyst"],
        "department": rng.choice(departments, 12),
        "location": rng.choice(["Delhi", "Bengaluru", "Mumbai", "Hyderabad"], 12),
        "openings": rng.integers(1, 4, 12),
        "candidates": rng.integers(8, 65, 12),
        "stage": rng.choice(["Sourcing", "Screening", "Interview", "Offer", "On Hold"], 12),
        "priority": rng.choice(["High", "Medium", "Low"], 12, p=[.3,.5,.2]),
    })
    resignations = pd.DataFrame({
        "exit_id": [f"EX-{i+1:03d}" for i in range(18)],
        "employee_id": rng.choice(employees.employee_id, 18, replace=False),
        "resignation_date": [today - pd.Timedelta(days=int(x)) for x in rng.integers(1, 365, 18)],
        "reason": rng.choice(["Career Growth", "Relocation", "Compensation", "Personal", "Higher Studies"], 18),
        "status": rng.choice(["Completed", "In Progress", "Pending"], 18),
    })
    departments_df = pd.DataFrame({
        "department_id": [f"D{i+1:02d}" for i in range(len(departments))],
        "department_name": departments,
        "headcount_target": [45, 15, 20, 30, 25, 25, 15, 15],
    })
    return {
        "Employees": employees, "Departments": departments_df, "Attendance": attendance,
        "Salaries": salaries, "Leave Requests": pd.DataFrame(leaves),
        "Recruitment": recruitment, "Resignations": resignations,
    }


@st.cache_data(show_spinner="Reading workbook…", max_entries=4)
def load_workbook_data(file_bytes: bytes):
    """Parse and normalize an uploaded workbook once per distinct file content."""
    book = pd.ExcelFile(BytesIO(file_bytes))
    return {name: clean_frame(pd.read_excel(book, sheet_name=name)) for name in book.sheet_names}


def classify_sheet(name):
    s = norm(name)
    if "employee" in s or "staff" in s: return "Employees"
    if "attend" in s or "timesheet" in s: return "Attendance"
    if "salar" in s or "payroll" in s or "compensation" in s: return "Payroll"
    if "leave" in s or "absence" in s: return "Leave Management"
    if "recruit" in s or "vacancy" in s or "hiring" in s: return "Recruitment"
    if "department" in s or "team" in s: return "Departments"
    if "resign" in s or "exit" in s or "termination" in s: return "Exits"
    return "Data Explorer"


def render_chart(frame, col, title, top=15):
    if not col or col not in frame.columns:
        st.info("This view needs a suitable column in the selected dataset.")
        return
    values = frame[col].fillna("Not specified").astype(str).value_counts().head(top)
    if values.empty:
        st.info("No values available for this chart.")
        return
    chart_df = values.rename_axis("Category").reset_index(name="Records")
    fig = px.bar(chart_df, x="Category", y="Records", title=title, text="Records")
    fig.update_layout(margin=dict(l=10, r=10, t=55, b=10), xaxis_title="", yaxis_title="Records")
    fig.update_traces(textposition="outside", cliponaxis=False)
    st.plotly_chart(fig, use_container_width=True)


def render_quality(frame):
    missing = frame.isna().sum()
    quality = pd.DataFrame({
        "Column": frame.columns,
        "Missing values": [int(missing[c]) for c in frame.columns],
        "Missing (%)": [round(100 * float(missing[c]) / max(len(frame), 1), 2) for c in frame.columns],
        "Unique values": [int(frame[c].nunique(dropna=True)) for c in frame.columns],
        "Data type": [str(frame[c].dtype) for c in frame.columns],
    })
    metric_row([
        ("Rows", f"{len(frame):,}", None),
        ("Duplicate rows", f"{int(frame.duplicated().sum()):,}", None),
        ("Missing cells", f"{int(missing.sum()):,}", None),
    ])
    st.dataframe(quality, use_container_width=True, hide_index=True)


# ---------------------------- Sidebar / data source -------------------------
with st.sidebar:
    st.markdown("## 👥 PeopleDesk")
    st.caption("Human Resources Management")
    st.markdown("---")
    uploaded = st.file_uploader("Upload HRMS workbook", type=["xlsx", "xls"], help="Your original workbook is read only; this app does not overwrite it.")
    st.markdown("#### Data source")
    source = st.radio("Choose data", ["Uploaded Excel workbook", "Fictional demo data"], index=1 if uploaded is None else 0)
    st.caption("Demo data is fictional and for demonstration only.")
    st.markdown("---")
    page = st.radio("WORKSPACE", [
        "Overview", "Employee Directory", "Attendance", "Payroll",
        "Leave Management", "Recruitment", "Exits & Attrition", "Data Explorer", "Data Quality"
    ], label_visibility="visible")
    st.markdown("---")
    st.caption("🔒 Uploaded workbook is used for this session only.")

# ---------------------------- Load data -------------------------------------
if source == "Uploaded Excel workbook":
    if uploaded is None:
        st.warning("Upload an Excel workbook in the sidebar, or select Fictional demo data.")
        st.stop()
    try:
        all_data = load_workbook_data(uploaded.getvalue())
    except Exception as exc:
        st.error(f"Could not read the workbook. Check that it is a valid Excel file. Details: {exc}")
        st.stop()
    if not all_data:
        st.error("No worksheets were found in this workbook.")
        st.stop()
    st.session_state["hrms_source_name"] = uploaded.name
else:
    all_data = demo_data()
    st.session_state["hrms_source_name"] = "Fictional demo data"

# Keep names as uploaded; map pages to matching sheets where possible.
def matching_frames(target):
    return [(name, frame) for name, frame in all_data.items() if classify_sheet(name) == target]

target_by_page = {
    "Employee Directory": "Employees", "Attendance": "Attendance", "Payroll": "Payroll",
    "Leave Management": "Leave Management", "Recruitment": "Recruitment",
    "Exits & Attrition": "Exits", "Data Explorer": "Data Explorer", "Data Quality": "Data Explorer",
}
target = target_by_page.get(page)
if target:
    candidates = matching_frames(target)
    selected_sheet = candidates[0][0] if candidates else st.selectbox(
        f"No dedicated {page} sheet was detected. Choose a worksheet to inspect.",
        list(all_data.keys(),), key=f"fallback_sheet_{page}"
    )
else:
    selected_sheet = next((n for n in all_data if classify_sheet(n) == "Employees"), list(all_data.keys())[0])
frame = all_data[selected_sheet].copy()

# Header
st.markdown(f"""
<div class="hero">
  <h1>👥 PeopleDesk <span style="font-weight:400">/ {page}</span></h1>
  <div class="muted">Comprehensive HR Management System · Source: {st.session_state['hrms_source_name']}</div>
</div>
""", unsafe_allow_html=True)
if source == "Fictional demo data":
    st.info("Demo mode: all names, employee IDs, salaries, attendance, leave and recruitment records shown here are synthetic examples.")

# Core inferred columns
employee_id_col = find_col(frame, ["employee_id", "emp_id", "staff_id", "employee code"], ["employeeid", "empid"])
name_col = find_col(frame, ["employee_name", "name", "full_name", "employee"], ["employeename", "fullname"])
dept_col = find_col(frame, ["department_name", "department", "dept_name", "team"], ["department", "dept"])
status_col = find_col(frame, ["status", "employment_status", "attendance_status", "payroll_status"], ["status"])
salary_col = find_col(frame, ["net_salary", "monthly_salary", "salary", "annual_salary", "gross_salary", "basic_salary"], ["salary", "wage", "pay"])
date_col = find_col(frame, ["joining_date", "attendance_date", "pay_date", "date", "resignation_date", "start_date"], ["date"])

# ---------------------------- Page renderers --------------------------------
if page == "Overview":
    # Build the overview from each relevant worksheet rather than the currently
    # selected sheet, so changing a sidebar page does not distort the KPIs.
    emp_pairs = matching_frames("Employees")
    emp_df = emp_pairs[0][1].copy() if emp_pairs else pd.DataFrame()
    att_pairs = matching_frames("Attendance")
    att_df = att_pairs[0][1].copy() if att_pairs else pd.DataFrame()
    pay_pairs = matching_frames("Payroll")
    pay_df = pay_pairs[0][1].copy() if pay_pairs else pd.DataFrame()
    leave_pairs = matching_frames("Leave Management")
    leave_df = leave_pairs[0][1].copy() if leave_pairs else pd.DataFrame()

    emp_status = find_col(emp_df, ["employment_status", "employee_status", "status"], ["employmentstatus", "employeestatus", "status"])
    emp_dept = find_col(emp_df, ["department_name", "department", "dept_name", "team"], ["department", "dept", "team"])
    join_col = find_col(emp_df, ["joining_date", "date_of_joining", "hire_date", "start_date"], ["joiningdate", "dateofjoining", "hiredate"])
    emp_salary = find_col(emp_df, ["annual_salary", "monthly_salary", "salary", "ctc", "gross_salary", "net_salary"], ["salary", "wage", "compensation", "ctc"])
    att_status = find_col(att_df, ["attendance_status", "status", "attendance"], ["attendancestatus", "status"])
    att_date = find_col(att_df, ["attendance_date", "date", "work_date"], ["attendancedate", "workdate", "date"])
    pay_salary = find_col(pay_df, ["net_salary", "gross_salary", "monthly_salary", "salary", "pay_amount", "amount"], ["salary", "payamount", "amount"])
    pay_period = find_col(pay_df, ["pay_month", "pay_period", "pay_date", "month"], ["paymonth", "payperiod", "paydate"])
    leave_status = find_col(leave_df, ["approval_status", "leave_status", "status"], ["approvalstatus", "leavestatus", "status"])

    if emp_df.empty:
        st.warning("No employee worksheet was detected. Upload a workbook with an employee/staff sheet to see workforce KPIs.")
    total_employees = len(emp_df)
    if emp_status:
        status_text = emp_df[emp_status].astype(str).str.strip().str.lower()
        active_count = int(status_text.isin(["active", "employed", "working", "permanent", "on leave"] ).sum())
    else:
        active_count = total_employees
    today = pd.Timestamp.today().normalize()
    if join_col:
        join_dates = safe_date_series(emp_df, join_col)
        new_hires = int(join_dates.between(today - pd.Timedelta(days=29), today, inclusive="both").sum())
    else:
        new_hires = 0
    if att_status:
        att_text = att_df[att_status].astype(str).str.strip().str.lower()
        valid_attendance = ~att_text.isin(["", "nan", "none", "not marked", "unknown"])
        attended = att_text.str.contains(r"present|work from home|\bwfh\b|half day|on time", regex=True, na=False)
        attendance_rate = (100 * int((attended & valid_attendance).sum()) / max(int(valid_attendance.sum()), 1))
        attendance_display = f"{attendance_rate:.1f}%" if valid_attendance.any() else "—"
    else:
        attendance_display = "—"
    if pay_salary:
        pay_values = pd.to_numeric(pay_df[pay_salary].astype(str).str.replace(r"[^0-9.\-]", "", regex=True), errors="coerce")
        if pay_period:
            pay_dates = safe_date_series(pay_df, pay_period)
            current_month = today.replace(day=1)
            current_pay = pay_values[pay_dates.dt.to_period("M") == current_month.to_period("M")]
            payroll_total = current_pay.sum() if current_pay.notna().any() else pay_values.sum()
            payroll_label = "Current-month payroll" if current_pay.notna().any() else "Payroll total"
        else:
            payroll_total = pay_values.sum()
            payroll_label = "Payroll total"
    else:
        payroll_total, payroll_label = None, "Payroll total"
    pending_leave = int(leave_df[leave_status].astype(str).str.strip().str.lower().eq("pending").sum()) if leave_status else 0

    metric_row([
        ("Total employees", f"{total_employees:,}", "Rows in the detected employee worksheet"),
        ("Active employees", f"{active_count:,}", "Based on employment status; on-leave employees are included in workforce headcount"),
        ("New hires · 30 days", f"{new_hires:,}" if join_col else "—", "Employees whose joining date falls in the last 30 days"),
        ("Attendance rate", attendance_display, "Share of marked attendance records recorded as present, WFH or half-day"),
        ("Pending leave", f"{pending_leave:,}" if leave_status else "—", "Leave requests awaiting a decision"),
    ])
    if payroll_total is not None:
        st.metric(payroll_label, money(payroll_total), help="Sum of the detected salary field; confirm whether the workbook stores monthly or annual amounts.")

    st.markdown("### Workforce analytics")
    left, right = st.columns(2)
    with left:
        if emp_dept and not emp_df.empty:
            dept_counts = emp_df[emp_dept].fillna("Not specified").astype(str).value_counts().rename_axis("Department").reset_index(name="Employees")
            fig = px.bar(dept_counts, x="Department", y="Employees", text="Employees", title="Headcount by department")
            fig.update_layout(margin=dict(l=10, r=10, t=55, b=10), xaxis_title="", yaxis_title="Employees")
            fig.update_traces(textposition="outside", cliponaxis=False)
            st.plotly_chart(fig, use_container_width=True)
        else:
            st.info("A department column is needed for the headcount chart.")
    with right:
        if emp_status and not emp_df.empty:
            status_counts = emp_df[emp_status].fillna("Not specified").astype(str).value_counts().rename_axis("Status").reset_index(name="Employees")
            fig = px.pie(status_counts, names="Status", values="Employees", hole=.55, title="Employment status mix")
            fig.update_layout(margin=dict(l=10, r=10, t=55, b=10), legend_title_text="Status")
            st.plotly_chart(fig, use_container_width=True)
        elif emp_salary and not emp_df.empty:
            nums = pd.to_numeric(emp_df[emp_salary], errors="coerce").dropna()
            if not nums.empty:
                fig = px.histogram(x=nums, nbins=18, title="Salary distribution")
                st.plotly_chart(fig, use_container_width=True)
            else:
                st.info("No employment status or numeric salary values were found.")
        else:
            st.info("An employment status column is needed for this chart.")

    left, right = st.columns(2)
    with left:
        if att_status and att_date and not att_df.empty:
            att_plot = att_df.copy()
            att_plot["_date"] = safe_date_series(att_plot, att_date)
            att_plot["_status"] = att_plot[att_status].astype(str).str.strip().str.title()
            att_plot = att_plot.dropna(subset=["_date"])
            if not att_plot.empty:
                daily = att_plot.groupby(["_date", "_status"]).size().reset_index(name="Records")
                daily = daily[daily["_date"] >= today - pd.Timedelta(days=29)]
                if not daily.empty:
                    fig = px.line(daily, x="_date", y="Records", color="_status", markers=True, title="Attendance trend · last 30 days")
                    fig.update_layout(margin=dict(l=10, r=10, t=55, b=10), xaxis_title="Date", yaxis_title="Records", legend_title_text="Status")
                    st.plotly_chart(fig, use_container_width=True)
                else:
                    st.info("No attendance records in the last 30 days.")
            else:
                st.info("No valid attendance dates were found.")
        else:
            st.info("Attendance trend requires both a date and a status column.")
    with right:
        if pay_salary and not pay_df.empty:
            pay_plot = pay_df.copy()
            pay_plot["_salary_value"] = pd.to_numeric(pay_plot[pay_salary].astype(str).str.replace(r"[^0-9.\-]", "", regex=True), errors="coerce")
            pay_dept = find_col(pay_plot, ["department", "department_name", "dept_name"], ["department", "dept"])
            if pay_dept and pay_plot["_salary_value"].notna().any():
                by_dept = pay_plot.groupby(pay_dept, dropna=False)["_salary_value"].sum().reset_index()
                by_dept[pay_dept] = by_dept[pay_dept].fillna("Not specified").astype(str)
                fig = px.bar(by_dept, x=pay_dept, y="_salary_value", title="Payroll by department", labels={"_salary_value": "Salary total", pay_dept: "Department"})
                fig.update_layout(margin=dict(l=10, r=10, t=55, b=10), xaxis_title="", yaxis_title="Amount")
                st.plotly_chart(fig, use_container_width=True)
            else:
                st.info("A department field is needed to compare payroll by department.")
        else:
            st.info("Upload a payroll worksheet with a numeric salary field to see payroll analytics.")

    st.markdown("### Recent employee records")
    if not emp_df.empty:
        recent = emp_df.copy()
        if join_col:
            recent["_sort_join_date"] = safe_date_series(recent, join_col)
            recent = recent.sort_values("_sort_join_date", ascending=False, na_position="last").drop(columns=["_sort_join_date"])
        st.dataframe(recent.head(10), use_container_width=True, hide_index=True)
        st.download_button("Download employee snapshot (CSV)", csv_bytes(emp_df),
                           file_name="employee_snapshot.csv", mime="text/csv")
    else:
        st.info("Employee records will appear here when an employee worksheet is available.")

elif page == "Employee Directory":
    st.subheader("Employee directory")
    search = st.text_input("Search by employee name, ID, department, title or email", placeholder="e.g. EMP1001 or Engineering")
    filters = st.columns(3)
    with filters[0]:
        if dept_col:
            depts = sorted(frame[dept_col].dropna().astype(str).unique().tolist())
            dept_filter = st.multiselect("Department", depts, default=[])
        else:
            dept_filter = []
    with filters[1]:
        if status_col:
            statuses = sorted(frame[status_col].dropna().astype(str).unique().tolist())
            status_filter = st.multiselect("Employment status", statuses, default=[])
        else:
            status_filter = []
    with filters[2]:
        location_col = find_col(frame, ["location", "office", "city"], ["location", "office", "city"])
        locations = sorted(frame[location_col].dropna().astype(str).unique().tolist()) if location_col else []
        location_filter = st.multiselect("Location", locations, default=[])
    shown = frame.copy()
    if search.strip():
        search_text = search.strip().casefold()
        mask = shown.astype(str).agg(" ".join, axis=1).str.casefold().str.contains(search_text, regex=False, na=False)
        shown = shown[mask]
    if dept_col and dept_filter:
        shown = shown[shown[dept_col].astype(str).isin(dept_filter)]
    if status_col and status_filter:
        shown = shown[shown[status_col].astype(str).isin(status_filter)]
    if location_col and location_filter:
        shown = shown[shown[location_col].astype(str).isin(location_filter)]
    st.caption(f"{len(shown):,} employees/records match the current filters.")
    st.dataframe(shown, use_container_width=True, hide_index=True)
    st.download_button("Export filtered directory as CSV", csv_bytes(shown),
                       file_name="employee_directory_filtered.csv", mime="text/csv")
    st.download_button("Export filtered directory as Excel", excel_bytes({"Employees": shown}),
                       file_name="employee_directory_filtered.xlsx",
                       mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")

elif page == "Attendance":
    st.subheader("Attendance & time tracking")
    if status_col:
        status_values = frame[status_col].astype(str).str.lower()
        present = int(status_values.str.contains("present|work from home|wfh|half day|on time", regex=True).sum())
        absent = int(status_values.str.contains("absent|missing", regex=True).sum())
        leave = int(status_values.str.contains("leave", regex=True).sum())
        metric_row([("Records", f"{len(frame):,}", None), ("Present / WFH / half-day", f"{present:,}", None),
                    ("Absent", f"{absent:,}", None), ("Leave marked", f"{leave:,}", None)])
        render_chart(frame, status_col, "Attendance status distribution")
    else:
        metric_row([("Attendance records", f"{len(frame):,}", None),
                    ("Columns available", f"{len(frame.columns):,}", None)])
        st.info("Could not identify a status column. The raw attendance records are still available below.")
    st.dataframe(frame, use_container_width=True, hide_index=True)
    st.download_button("Download attendance CSV", csv_bytes(frame),
                       file_name="attendance.csv", mime="text/csv")

elif page == "Payroll":
    st.subheader("Payroll overview")
    if salary_col and pd.to_numeric(frame[salary_col], errors="coerce").notna().any():
        nums = pd.to_numeric(frame[salary_col], errors="coerce")
        metric_row([("Payroll records", f"{len(frame):,}", None),
                    (f"Total {salary_col.replace('_',' ').title()}", money(nums.sum()), "Sum of the detected salary field"),
                    ("Average value", money(nums.mean()), "Mean of available numeric salary values"),
                    ("Missing salary values", f"{int(nums.isna().sum()):,}", None)])
        valid = frame.assign(_salary=nums).dropna(subset=["_salary"])
        if not valid.empty:
            fig = px.histogram(valid, x="_salary", nbins=20, title="Salary distribution")
            fig.update_layout(xaxis_title=salary_col.replace("_", " ").title(), yaxis_title="Employees / records")
            st.plotly_chart(fig, use_container_width=True)
    else:
        st.info("No salary-like numeric field was detected on this worksheet. Check the workbook sheet selector or upload a payroll sheet.")
        metric_row([("Payroll records", f"{len(frame):,}", None), ("Columns", f"{len(frame.columns):,}", None)])
    st.caption("Payroll figures are shown for analysis only. This demo does not calculate statutory deductions or process payments.")
    st.dataframe(frame, use_container_width=True, hide_index=True)
    st.download_button("Download payroll CSV", csv_bytes(frame),
                       file_name="payroll_export.csv", mime="text/csv")

elif page == "Leave Management":
    st.subheader("Leave request register")
    if status_col:
        pending_mask = frame[status_col].astype(str).str.lower().eq("pending")
        metric_row([("Leave records", f"{len(frame):,}", None),
                    ("Pending", f"{int(pending_mask.sum()):,}", None),
                    ("Approved", f"{int(frame[status_col].astype(str).str.lower().eq('approved').sum()):,}", None),
                    ("Rejected", f"{int(frame[status_col].astype(str).str.lower().eq('rejected').sum()):,}", None)])
        choices = sorted(frame[status_col].dropna().astype(str).unique().tolist())
        chosen_status = st.multiselect("Filter by request status", choices, default=[])
        shown = frame[frame[status_col].astype(str).isin(chosen_status)] if chosen_status else frame
    else:
        shown = frame
        st.info("No leave status column was detected. Showing all available records.")
    st.dataframe(shown, use_container_width=True, hide_index=True)
    st.caption("This is a reporting view. Approval changes are not written back to your original workbook.")
    st.download_button("Download leave register CSV", csv_bytes(shown),
                       file_name="leave_register.csv", mime="text/csv")

elif page == "Recruitment":
    st.subheader("Recruitment pipeline")
    stage_col = find_col(frame, ["stage", "pipeline_stage", "application_status", "status"], ["stage", "status"])
    job_col = find_col(frame, ["job_title", "position", "role", "vacancy"], ["jobtitle", "position", "vacancy"])
    candidate_col = find_col(frame, ["candidates", "candidate_count", "applicants"], ["candidate", "applicant"])
    if stage_col:
        render_chart(frame, stage_col, "Open requisitions by stage")
        metric_row([("Requisitions", f"{len(frame):,}", None),
                    ("Pipeline stages", f"{frame[stage_col].nunique(dropna=True):,}", None),
                    ("High priority", f"{int(frame[find_col(frame, ['priority'], ['priority'])].astype(str).str.lower().eq('high').sum()):,}" if find_col(frame, ['priority'], ['priority']) else "—", None)])
    if candidate_col and pd.to_numeric(frame[candidate_col], errors="coerce").notna().any():
        metric_row([("Candidates", f"{int(pd.to_numeric(frame[candidate_col], errors='coerce').sum()):,}", "Sum of candidate counts across requisitions")])
    st.dataframe(frame, use_container_width=True, hide_index=True)
    st.download_button("Download recruitment pipeline CSV", csv_bytes(frame),
                       file_name="recruitment_pipeline.csv", mime="text/csv")

elif page == "Exits & Attrition":
    st.subheader("Employee exits & attrition")
    reason_col = find_col(frame, ["reason", "exit_reason", "resignation_reason"], ["reason"])
    exit_date_col = find_col(frame, ["resignation_date", "exit_date", "last_working_day", "date"], ["resignationdate", "exitdate"])
    metric_row([("Exit records", f"{len(frame):,}", None),
                ("Unique employees", f"{frame[employee_id_col].nunique(dropna=True):,}" if employee_id_col else "—", None),
                ("Exit reasons", f"{frame[reason_col].nunique(dropna=True):,}" if reason_col else "—", None)])
    left, right = st.columns(2)
    with left:
        if reason_col:
            render_chart(frame, reason_col, "Exit reasons")
        else:
            st.info("No exit reason field found.")
    with right:
        if exit_date_col:
            dates = safe_date_series(frame, exit_date_col).dropna()
            if not dates.empty:
                by_month = dates.dt.to_period("M").astype(str).value_counts().sort_index().rename_axis("Month").reset_index(name="Exits")
                fig = px.line(by_month, x="Month", y="Exits", markers=True, title="Exits by month")
                st.plotly_chart(fig, use_container_width=True)
            else:
                st.info("No valid exit dates found.")
        else:
            st.info("No exit date field found.")
    st.dataframe(frame, use_container_width=True, hide_index=True)
    st.download_button("Download exit records CSV", csv_bytes(frame),
                       file_name="exit_records.csv", mime="text/csv")

elif page == "Data Quality":
    st.subheader(f"Data quality report — {selected_sheet}")
    render_quality(frame)
    st.markdown("### Validation checks")
    duplicates = frame[frame.duplicated(keep=False)]
    email_col = find_col(frame, ["email", "email_address", "work_email"], ["email"])
    id_col = find_col(frame, ["employee_id", "emp_id", "staff_id", "employee_code"], ["employeeid", "empid", "staffid"])
    validation_rows = []
    validation_rows.append({"Check": "Exact duplicate rows", "Issue count": int(frame.duplicated().sum()), "Result": "Review" if frame.duplicated().any() else "Pass"})
    if id_col:
        duplicate_ids = frame[id_col].dropna().astype(str).duplicated(keep=False)
        validation_rows.append({"Check": f"Duplicate values in {id_col}", "Issue count": int(duplicate_ids.sum()), "Result": "Review" if duplicate_ids.any() else "Pass"})
    if email_col:
        emails = frame[email_col].dropna().astype(str).str.strip()
        invalid_email = ~emails.str.match(r"^[^@\s]+@[^@\s]+\.[^@\s]+$")
        duplicate_email = emails.str.lower().duplicated(keep=False)
        validation_rows.append({"Check": "Invalid email formats", "Issue count": int(invalid_email.sum()), "Result": "Review" if invalid_email.any() else "Pass"})
        validation_rows.append({"Check": "Duplicate email addresses", "Issue count": int(duplicate_email.sum()), "Result": "Review" if duplicate_email.any() else "Pass"})
    st.dataframe(pd.DataFrame(validation_rows), use_container_width=True, hide_index=True)
    if duplicates.empty:
        st.success("No exact duplicate rows found in this worksheet.")
    else:
        st.warning(f"{len(duplicates):,} rows belong to exact-duplicate groups.")
        st.dataframe(duplicates, use_container_width=True, hide_index=True)
    st.download_button("Download data quality report (Excel)", excel_bytes({"Quality": pd.DataFrame({
        "Column": frame.columns,
        "Missing values": [int(frame[c].isna().sum()) for c in frame.columns],
        "Unique values": [int(frame[c].nunique(dropna=True)) for c in frame.columns],
        "Data type": [str(frame[c].dtype) for c in frame.columns],
    }), "Data": frame}), file_name="hrms_data_quality.xlsx",
       mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")

else:  # Data Explorer
    st.subheader("Explore any worksheet")
    selected = st.selectbox("Worksheet", list(all_data.keys()), index=list(all_data.keys()).index(selected_sheet))
    frame = all_data[selected].copy()
    st.caption(f"{len(frame):,} rows · {len(frame.columns):,} columns")
    filter_cols = st.multiselect("Choose columns to filter", list(frame.columns), default=[], key=f"explorer_cols_{selected}")
    filtered = frame.copy()
    for i, col in enumerate(filter_cols):
        values = sorted(filtered[col].dropna().astype(str).unique().tolist())
        if len(values) > 250:
            st.caption(f"Skipping {col}: too many distinct values ({len(values):,}).")
            continue
        selected_values = st.multiselect(
            f"Values for {col}", options=values, default=[],
            key=f"explorer_values_{selected}_{col}",
            help="Leave empty to include all values; select one or more values to narrow the table.",
        )
        if selected_values:
            filtered = filtered[filtered[col].astype(str).isin(selected_values)]
    st.caption(f"Showing {len(filtered):,} of {len(frame):,} rows.")
    st.dataframe(filtered, use_container_width=True, hide_index=True)
    c1, c2 = st.columns(2)
    with c1:
        st.download_button("Export filtered data as CSV", csv_bytes(filtered),
                           file_name=f"{norm(selected)}_filtered.csv", mime="text/csv", use_container_width=True)
    with c2:
        st.download_button("Export filtered data as Excel", excel_bytes({selected: filtered}),
                           file_name=f"{norm(selected)}_filtered.xlsx",
                           mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                           use_container_width=True)
    st.download_button("Export all worksheets as Excel", excel_bytes(all_data),
                       file_name="peopledesk_all_worksheets.xlsx",
                       mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")

st.markdown("---")
st.caption("PeopleDesk HRMS · Built with Streamlit · Demo data is fictional; uploaded workbooks are not overwritten.")

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
div[data-testid="stDataFrame"] {border: 1px solid rgba(128,128,128,.15); border-radius: 10px;}
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


def excel_bytes(sheets: dict[str, pd.DataFrame]) -> bytes:
    buff = BytesIO()
    with pd.ExcelWriter(buff, engine="openpyxl") as writer:
        for name, frame in sheets.items():
            frame.to_excel(writer, sheet_name=str(name)[:31] or "Data", index=False)
    return buff.getvalue()


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
    join_start = pd.Timestamp("2019-01-01")
    join_days = rng.integers(0, 7 * 365, n)
    employees = pd.DataFrame({
        "employee_id": [f"EMP{i+1001}" for i in range(n)],
        "employee_name": names,
        "email": [f"{norm(x)}{i%9}@example.com" for i, x in enumerate(names)],
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
    salaries = employees[["employee_id", "employee_name", "department", "annual_salary"]].copy()
    salaries["pay_month"] = today.replace(day=1)
    salaries["basic_salary"] = (salaries["annual_salary"] / 12 * .5).round(0)
    salaries["allowances"] = (salaries["annual_salary"] / 12 * .35).round(0)
    salaries["deductions"] = (salaries["annual_salary"] / 12 * .08).round(0)
    salaries["net_salary"] = (salaries["annual_salary"] / 12 - salaries["deductions"]).round(0)
    salaries["payroll_status"] = rng.choice(["Processed", "Processed", "Pending"], n, p=[.92,.05,.03])
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
        book = pd.ExcelFile(uploaded)
        all_data = {name: clean_frame(pd.read_excel(book, sheet_name=name)) for name in book.sheet_names}
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
    emp_pairs = matching_frames("Employees")
    emp_df = emp_pairs[0][1].copy() if emp_pairs else frame.copy()
    emp_status = find_col(emp_df, ["status", "employment_status"], ["status"])
    emp_dept = find_col(emp_df, ["department", "department_name", "dept_name"], ["department", "dept"])
    emp_salary = find_col(emp_df, ["annual_salary", "salary", "monthly_salary", "net_salary"], ["salary", "pay"])
    active_count = int(emp_df[emp_df[emp_status].astype(str).str.lower().isin(["active", "employed", "working", "permanent"])].shape[0]) if emp_status else len(emp_df)
    att_pairs = matching_frames("Attendance")
    att_df = att_pairs[0][1] if att_pairs else pd.DataFrame()
    att_status = find_col(att_df, ["attendance_status", "status"], ["status"])
    present_count = int(att_df[att_status].astype(str).str.lower().str.contains("present|work from home|wfh|half day", regex=True, na=False).sum()) if att_status else 0
    pay_pairs = matching_frames("Payroll")
    pay_df = pay_pairs[0][1] if pay_pairs else pd.DataFrame()
    pay_salary = find_col(pay_df, ["net_salary", "monthly_salary", "salary", "gross_salary"], ["salary", "pay"])
    total_pay = pd.to_numeric(pay_df[pay_salary], errors="coerce").sum() if pay_salary else 0
    leave_pairs = matching_frames("Leave Management")
    leave_df = leave_pairs[0][1] if leave_pairs else pd.DataFrame()
    leave_status = find_col(leave_df, ["status", "approval_status"], ["status"])
    pending_leave = int(leave_df[leave_status].astype(str).str.lower().eq("pending").sum()) if leave_status else 0
    metric_row([
        ("Total employees", f"{len(emp_df):,}", "Records in the employee worksheet"),
        ("Active employees", f"{active_count:,}", "Based on employment status when available"),
        ("Attendance marked present", f"{present_count:,}", "Present, WFH and half-day records in selected attendance data"),
        ("Pending leave requests", f"{pending_leave:,}", "Requests whose status is Pending"),
    ])
    st.markdown("### Workforce snapshot")
    left, right = st.columns(2)
    with left:
        if emp_dept:
            render_chart(emp_df, emp_dept, "Employees by department")
        else:
            st.info("No department column detected in the employee sheet.")
    with right:
        if emp_status:
            render_chart(emp_df, emp_status, "Employment status")
        elif emp_salary:
            numeric = pd.to_numeric(emp_df[emp_salary], errors="coerce").dropna()
            if not numeric.empty:
                fig = px.histogram(numeric, nbins=18, title=f"Distribution of {emp_salary.replace('_',' ').title()}")
                st.plotly_chart(fig, use_container_width=True)
        else:
            st.info("Add a status, department or salary field for more workforce charts.")
    st.markdown("### Recent employee records")
    st.dataframe(emp_df.head(10), use_container_width=True, hide_index=True)
    st.download_button("Download employee snapshot (CSV)", emp_df.to_csv(index=False).encode("utf-8-sig"),
                       file_name="employee_snapshot.csv", mime="text/csv")

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
        mask = shown.astype(str).apply(lambda s: s.str.contains(search.strip(), case=False, na=False, regex=False)).any(axis=1)
        shown = shown[mask]
    if dept_col and dept_filter:
        shown = shown[shown[dept_col].astype(str).isin(dept_filter)]
    if status_col and status_filter:
        shown = shown[shown[status_col].astype(str).isin(status_filter)]
    if location_col and location_filter:
        shown = shown[shown[location_col].astype(str).isin(location_filter)]
    st.caption(f"{len(shown):,} employees/records match the current filters.")
    st.dataframe(shown, use_container_width=True, hide_index=True)
    st.download_button("Export filtered directory as CSV", shown.to_csv(index=False).encode("utf-8-sig"),
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
    st.download_button("Download attendance CSV", frame.to_csv(index=False).encode("utf-8-sig"),
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
    st.download_button("Download payroll CSV", frame.to_csv(index=False).encode("utf-8-sig"),
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
    st.download_button("Download leave register CSV", shown.to_csv(index=False).encode("utf-8-sig"),
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
    st.download_button("Download recruitment pipeline CSV", frame.to_csv(index=False).encode("utf-8-sig"),
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
    st.download_button("Download exit records CSV", frame.to_csv(index=False).encode("utf-8-sig"),
                       file_name="exit_records.csv", mime="text/csv")

elif page == "Data Quality":
    st.subheader(f"Data quality report — {selected_sheet}")
    render_quality(frame)
    st.markdown("### Duplicate records")
    duplicates = frame[frame.duplicated(keep=False)]
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
            f"Values for {col}", options=values, default=values,
            key=f"explorer_values_{selected}_{col}",
        )
        if selected_values:
            filtered = filtered[filtered[col].astype(str).isin(selected_values)]
        else:
            filtered = filtered.iloc[0:0]
    st.caption(f"Showing {len(filtered):,} of {len(frame):,} rows.")
    st.dataframe(filtered, use_container_width=True, hide_index=True)
    c1, c2 = st.columns(2)
    with c1:
        st.download_button("Export filtered data as CSV", filtered.to_csv(index=False).encode("utf-8-sig"),
                           file_name=f"{norm(selected)}_filtered.csv", mime="text/csv", use_container_width=True)
    with c2:
        st.download_button("Export filtered data as Excel", excel_bytes({selected: filtered}),
                           file_name=f"{norm(selected)}_filtered.xlsx",
                           mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                           use_container_width=True)

st.markdown("---")
st.caption("PeopleDesk HRMS · Built with Streamlit · Demo data is fictional; uploaded workbooks are not overwritten.")

Comprehensive HR Management System (HRMS)
A Python-based Human Resource Management dashboard built with Streamlit. The application helps users explore HR data stored in Excel workbooks through summary metrics, interactive visualizations, dedicated HR data views, and data-quality checks.
Project Overview
Comprehensive HR Management brings related HR spreadsheet data into one dashboard. Users can upload an Excel workbook or explore the application's clearly labelled fictional demo dataset. The dashboard displays information based on the worksheets and columns available in the selected data.

The project is intended for HR data exploration, academic demonstration, and portfolio use. It should not be treated as a replacement for a production HR platform or as a system of record.
Objectives
- Provide a single dashboard for exploring HR information stored in Excel.
- Summarize workforce indicators when the required data fields are available.
- Organize common HR data into focused views.
- Visualize available data using interactive charts.
- Help users inspect missing values, duplicate rows, column types, and other basic data-quality indicators.
- Allow supported tables or data to be exported for further analysis.
Features
- HR Overview Dashboard — displays available workforce summary metrics.
- Workforce Analytics — charts such as department headcount and employment-status distribution when suitable fields are present.
- Employee Directory — inspect employee-related worksheet records.
- Attendance View — explore attendance data and related fields when available.
- Payroll View — inspect payroll or salary-related records when available.
- Leave Management View — review leave records and pending requests when the relevant data is present.
- Recruitment View — inspect recruitment-related worksheet records.
- Exits & Attrition View — explore exit- or termination-related records when available.
- Data Explorer — select and inspect worksheet data.
- Data Quality Checks — review row counts, duplicate rows, missing cells, missing-value percentages, unique values, and column data types.
- Excel Upload — load an .xlsx or .xls workbook.
- Demo Dataset — explore clearly labelled fictional sample data.
- Data Export — download supported data in CSV or Excel format.
Note: The dashboard's displayed metrics and charts depend on the workbook's worksheets and column names. A view may show limited information when the corresponding data is not available.

Technologies Used
- Python — application logic and data processing
- Streamlit — dashboard interface
- Pandas — tabular data handling and analysis
- NumPy — fictional demonstration data generation
- Plotly Express — interactive charts
- OpenPyXL — Excel workbook support and export
Dependency versions are defined in (requirements.txt).
Getting Started
Prerequisites
- Python installed on your computer
- Git (optional, for cloning the repository)
- The project files, including app.py and requirements.txt
1. Clone the repository
Replace the repository URL below if your GitHub repository address changes.
git clone https://github.com/maanbeniwal448-bit/HRMS.git
cd HRMS
Alternatively, download the repository as a ZIP file from GitHub and extract it.
2. Create a virtual environment (recommended)
Windows
python -m venv venv
venv\Scripts\activate
macOS / Linux
python3 -m venv venv
source venv/bin/activate
3. Install dependencies
python -m pip install -r requirements.txt
4. Run the Streamlit application
python -m streamlit run app.py
Streamlit will show a local address in the terminal, usually:
http://localhost:8501
Open the displayed address in your browser.
How to Use
1. Start the application using the command above.
2. Choose the available demo dataset or upload an HR Excel workbook using the sidebar controls.
3. Open the dashboard section you want to explore.
4. Review the available summary metrics, tables, charts, and data-quality information.
5. If the selected view supports it, download the displayed data for further analysis.
Working with Excel Data
- Use a workbook containing the HR worksheets relevant to your analysis.
- Available views and calculations depend on the worksheet names and column names in the workbook.
- Check the Data Explorer or Data Quality section when you need to inspect the structure and completeness of your data.
- Keep a backup of your original workbook. The application is intended to read the uploaded workbook for the current session rather than overwrite the source file.
Project Structure
The core files used to run the application are:
HRMS/
├── app.py              # Streamlit dashboard application
├── requirements.txt    # Python package dependencies
└── README.md           # Project documentation
Your repository may contain additional files, such as the HR Excel dataset or project materials.
Data and Privacy Notes
- Only upload files you are authorized to use.
- Avoid uploading real employee personal, payroll, or other confidential information to public repositories.
- Do not commit sensitive datasets, credentials, passwords, or secret configuration files to GitHub.
- The demo dataset is intended for demonstration and should not be interpreted as real employee information.
- Review any exported file before sharing it.
Limitations
- The information displayed depends on the data supplied by the user.
- Metrics and visualizations may not be available if the expected data fields are missing or named differently.
- The project focuses on exploring and analyzing HR data; features not implemented in the application should not be assumed to be available.
- Organizational results or performance improvements have not been measured or specified here.
Future Improvements
Possible future enhancements, subject to implementation, include:
- More flexible mapping of different Excel column names.
- Additional configurable HR reports and filters.
- Expanded validation and data-quality reporting.
- Further usability improvements based on user feedback.
These are ideas for future development and are not claims about existing functionality.
Project Information
- Project Title: Comprehensive HR Management
- Project Type: HR data dashboard and analysis project
- Interface: Streamlit web application
- Repository: https://github.com/maanbeniwal448-bit/HRMS

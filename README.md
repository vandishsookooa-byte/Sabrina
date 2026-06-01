# Sabrina Attendance Dashboard (Flask Template)

This repository now includes a Python Flask template app that loads an attendance Excel file and renders a management dashboard for:

- OT1 / OT2 / OT3 monthly percentages by department
- Leave %, Missing Hours %, Attendance %, Absenteeism % (monthly, by department)
- Top employees by absenteeism, missing hours, and overtime
- Department rankings (attendance, absenteeism, missing hours, overtime)
- Leave utilization by leave type and department
- Absenteeism vs overtime trend/correlation by department

## 1) Install

```bash
python -m venv .venv
source .venv/bin/activate  # Windows: .venv\Scripts\activate
pip install -r requirements.txt
```

## 2) Configure Excel file

Set an environment variable with the file path:

```bash
export EXCEL_FILE_PATH="/absolute/path/to/Last 1 year attendance.xlsx"
```

Windows PowerShell example:

```powershell
$env:EXCEL_FILE_PATH="C:\Users\transport\Desktop\SABRINA ATTENDANCE\Last 1 year attendance.xlsx"
```

## 3) Run Flask server

```bash
python /tmp/workspace/vandishsookooa-byte/Sabrina/app.py
```

Optional port override:

```bash
PORT=5000 python /tmp/workspace/vandishsookooa-byte/Sabrina/app.py
```

## 4) Open dashboard

Visit:

`http://localhost:5000`

## Notes

- The app normalizes common attendance column names automatically.
- If your sheet uses different headers, update `COLUMN_ALIASES` in `app.py`.
- Interactive filters are available for date range, department, employee, and leave type.
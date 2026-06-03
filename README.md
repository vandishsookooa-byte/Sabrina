# Sabrina – Attendance Intelligence Dashboard

A comprehensive **Flask-based workforce analytics dashboard** providing real-time attendance KPIs, overtime analysis, leave utilisation, manpower risk assessment, and predictive forecasting — all presented in a **dark professional management reporting theme**.

---

## Features

| Category | Highlights |
|---|---|
| **KPI Cards** | Attendance %, Absenteeism %, OT %, Leave %, Missing Hrs %, Headcount with traffic-light indicators |
| **Monthly Trends** | Line charts for all metrics across the year |
| **Department Charts** | Bar charts comparing departments across all KPIs |
| **Heat Maps** | Monthly performance grids (Attendance, Absenteeism, Missing Hrs, OT) |
| **OT Analysis** | OT1 / OT2 / OT3 breakdown by department, top-5 rankings, monthly trends |
| **Leave Analysis** | Leave utilisation by type and department with donut + bar charts |
| **Employee Rankings** | Top 10 by Overtime, Missing Hours, and Absenteeism (using ID Card) |
| **Exception Reports** | Employees with excessive absenteeism, OT, missing hrs, no leave, or 3+ consecutive absences |
| **Manpower Risk** | Departments with increasing trends, low attendance streaks, or requiring intervention |
| **Forecasting** | Linear trend extrapolation for next 3 months (Attendance, Absenteeism, OT, Missing Hrs) |
| **Drill-Down** | Detailed department × employee analysis page |
| **Management Insights** | Automated recommendations and action points |
| **Filters** | Multi-select: Department, Employee, Month, Year, Leave Type, Date Range |

---

## Quick Start

### 1. Install dependencies

```bash
pip install -r requirements.txt
```

### 2. Set the Excel file path (optional)

```bash
# Linux / macOS
export EXCEL_FILE_PATH="/path/to/Last 1 year attendance.xlsx"

# Windows
set EXCEL_FILE_PATH=C:\Users\transport\Desktop\SABRINA ATTENDANCE\Last 1 year attendance.xlsx
```

If the file is not found or not set, the app generates **realistic sample data** automatically so the dashboard is fully functional for demonstration.

### 3. Run the server

```bash
python app.py
```

Then open **http://localhost:5000** in your browser.

### Optional environment variables

| Variable | Default | Description |
|---|---|---|
| `EXCEL_FILE_PATH` | Windows path in app.py | Path to the attendance Excel file |
| `PORT` | `5000` | Server port |
| `FLASK_DEBUG` | `0` | Set to `1` to enable debug mode |

---

## Expected Excel Columns

The app auto-detects common column name variants. The table below shows the internal name and accepted aliases:

| Internal Name | Accepted Column Headers |
|---|---|
| `id_card` | ID Card, Employee ID, EmpID, Card No |
| `department` | Department, Dept, Department Name |
| `date` | Date, Attendance Date, Att Date |
| `status` | Status, Attendance Status (P/A/L) |
| `hours_worked` | Hours Worked, Worked Hours, Working Hours |
| `leave_type` | Leave Type, LeaveType, Leave Category |
| `leave_quantity` | Leave Quantity, Leave Days, LeaveQty |
| `missing_hours` | Missing Hours, MissingHours, Late Hours |
| `ot1` | OT1, OT 1, Overtime 1 |
| `ot2` | OT2, OT 2, Overtime 2 |
| `ot3` | OT3, OT 3, Overtime 3 |

---

## Routes

| Route | Description |
|---|---|
| `/` | Executive Dashboard – KPIs, trends, rankings, risk analysis |
| `/drilldown` | Drill-Down Analytics – department × employee level detail |
| `/api/dashboard` | JSON API for executive dashboard data |
| `/api/drilldown` | JSON API for drill-down data |
| `/api/filters` | JSON API returning available filter options |
| `/api/export` | Export filtered attendance records to Excel |
| `/methodology` | Calculation methodology and threshold reference page |

All API endpoints accept these query parameters:
`departments`, `years`, `months`, `id_cards`, `leave_types`, `date_from`, `date_to`

---

## KPI Traffic Light Thresholds

| Metric | 🟢 Green | 🟡 Yellow | 🔴 Red |
|---|---|---|---|
| Attendance % | ≥ 95% | 85–95% | < 85% |
| Absenteeism % | ≤ 5% | 5–15% | > 15% |
| Leave % | ≤ 8% | 8–15% | > 15% |
| Missing Hours % | ≤ 5% | 5–10% | > 10% |
| Overtime % | ≤ 10% | 10–20% | > 20% |

---

## Project Structure

```
Sabrina/
├── app.py                  # Flask application & all data processing logic
├── requirements.txt        # Python dependencies
├── README.md
├── templates/
│   ├── base.html           # Dark theme base layout (nav, CDN links)
│   ├── index.html          # Executive Dashboard page
│   └── drilldown.html      # Drill-Down Analytics page
└── static/
    └── css/
        └── style.css       # Dark professional management theme
```

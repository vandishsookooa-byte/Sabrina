# Sabrina

Flask dashboard template for attendance analytics.

## Run locally

```bash
python -m venv .venv
source .venv/bin/activate  # On Windows: .venv\\Scripts\\activate
pip install -r requirements.txt
python app.py
```

Open `http://localhost:5000` and upload the attendance Excel file.

## What the dashboard includes

- Monthly department % for OT1, OT2, OT3
- Monthly leave %, missing hours %, attendance %, absenteeism %
- Top employees by absenteeism, missing hours, and overtime
- Overtime utilization and absenteeism-vs-overtime trend table
- Leave utilization by leave type and department
- Department rankings (attendance, absenteeism, missing hours, overtime)
- Interactive filters for date range, department, employee, leave type
- Executive summary cards and department attention highlights

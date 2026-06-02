"""
Sabrina Attendance Dashboard – Flask Application
================================================
Run:  python app.py
Env:  EXCEL_FILE_PATH  – path to the source Excel file (optional)
      PORT             – server port (default 5000)
      FLASK_DEBUG      – set to 1 for debug mode
"""

import os
import json
import random
import calendar
from io import BytesIO
from datetime import datetime, timedelta, date

import numpy as np
import pandas as pd
from flask import Flask, render_template, jsonify, request, send_file

# ─────────────────────────────────────────────────────────────────────────────
# Configuration
# ─────────────────────────────────────────────────────────────────────────────
EXCEL_FILE_PATH = os.getenv(
    "EXCEL_FILE_PATH",
    r"C:\Users\transport\Desktop\SABRINA ATTENDANCE\Last 1 year attendance.xlsx",
)
PORT = int(os.getenv("PORT", 5000))
DEBUG = os.getenv("FLASK_DEBUG", "0") == "1"

# KPI thresholds (for traffic-light indicators)
THRESHOLDS = {
    "att_pct":     {"green": 95, "yellow": 85},   # ≥green=good
    "abs_pct":     {"yellow": 5, "red": 15},       # ≤yellow=good, ≥red=bad
    "leave_pct":   {"yellow": 8, "red": 15},
    "missing_pct": {"yellow": 5, "red": 10},
    "ot_pct":      {"yellow": 10, "red": 20},
}

MONTH_NAMES = [
    "", "Jan", "Feb", "Mar", "Apr", "May", "Jun",
    "Jul", "Aug", "Sep", "Oct", "Nov", "Dec",
]

# ─────────────────────────────────────────────────────────────────────────────
# Column aliases – maps internal names → possible Excel column names
# ─────────────────────────────────────────────────────────────────────────────
COLUMN_ALIASES = {
    "id_card":        ["ID Card", "Employee ID", "EmpID", "ID", "Card No", "Emp ID", "Employee Id"],
    "employee_name":  ["Employee Name", "Name", "Emp Name", "Full Name", "Employee"],
    "department":     ["Department", "Dept", "Department Name", "DEPARTMENT"],
    "date":           ["Date", "Attendance Date", "Att Date", "Work Date", "ATTENDANCE_DATE"],
    "status":         ["Status", "Attendance Status", "Att Status", "ATTENDANCE"],
    "leave_type":     ["Leave Type", "LeaveType", "Leave Category", "LEAVE_TYPE"],
    "leave_quantity": ["Leave Quantity", "Leave Days", "LeaveQty", "Leave Qty", "No. of Days", "LEAVE_QTY"],
    "missing_hours":  ["Missing Hours", "MissingHours", "Late Hours", "Miss Hours", "MISSING_HOURS"],
    "ot1":            ["OT1", "OT 1", "Overtime 1", "OT-1", "Over Time 1", "OT_1"],
    "ot2":            ["OT2", "OT 2", "Overtime 2", "OT-2", "Over Time 2", "OT_2"],
    "ot3":            ["OT3", "OT 3", "Overtime 3", "OT-3", "Over Time 3", "OT_3"],
}

PRESENT_VALUES = {"P", "PRESENT", "A-P", "AP", "1"}
ABSENT_VALUES  = {
    "A", "ABSENT", "AB", "ABS", "0",
    "NA", "UA",
    "NOTIFIED ABSENCE", "UNNOTIFIED ABSENCE",
    "N/A", "U/A",
}
LEAVE_VALUES   = {"L", "LEAVE", "AL", "SL", "EL", "ML", "PL", "CL"}
APPROVED_LEAVE_TYPES = {
    "Annual Leave",
    "Sick Leave",
    "Emergency Leave",
    "Maternity Leave",
    "Unpaid Leave",
}

# ─────────────────────────────────────────────────────────────────────────────
# Sample data generation (used when the Excel file is not found)
# ─────────────────────────────────────────────────────────────────────────────
DEPARTMENTS = {
    "Dyeing":          {"size": 50, "att": 0.88, "ot1": 0.12, "ot2": 0.08, "ot3": 0.05, "miss": 0.06},
    "Knitting":        {"size": 35, "att": 0.90, "ot1": 0.10, "ot2": 0.07, "ot3": 0.04, "miss": 0.05},
    "Finishing":       {"size": 40, "att": 0.95, "ot1": 0.07, "ot2": 0.04, "ot3": 0.02, "miss": 0.03},
    "Quality Control": {"size": 20, "att": 0.93, "ot1": 0.06, "ot2": 0.03, "ot3": 0.01, "miss": 0.07},
    "Maintenance":     {"size": 15, "att": 0.91, "ot1": 0.18, "ot2": 0.12, "ot3": 0.06, "miss": 0.04},
    "Administration":  {"size": 25, "att": 0.97, "ot1": 0.03, "ot2": 0.02, "ot3": 0.01, "miss": 0.02},
    "Human Resources": {"size": 10, "att": 0.96, "ot1": 0.02, "ot2": 0.01, "ot3": 0.01, "miss": 0.01},
    "Finance":         {"size": 10, "att": 0.96, "ot1": 0.04, "ot2": 0.02, "ot3": 0.01, "miss": 0.02},
    "Logistics":       {"size": 20, "att": 0.89, "ot1": 0.09, "ot2": 0.06, "ot3": 0.03, "miss": 0.05},
    "Production":      {"size": 30, "att": 0.92, "ot1": 0.14, "ot2": 0.09, "ot3": 0.05, "miss": 0.04},
}

LEAVE_TYPES = ["Annual Leave", "Sick Leave", "Emergency Leave", "Maternity Leave", "Unpaid Leave"]


def _working_days_in_month(year: int, month: int) -> list:
    """Return a list of Mon-Sat dates for the given month/year."""
    _, last = calendar.monthrange(year, month)
    return [
        date(year, month, d)
        for d in range(1, last + 1)
        if date(year, month, d).weekday() < 6  # 0=Mon, 5=Sat, 6=Sun
    ]


def generate_sample_data() -> pd.DataFrame:
    """Generate one year of realistic attendance data."""
    random.seed(42)
    np.random.seed(42)

    current_year = datetime.now().year
    year = current_year - 1  # last full year
    months = range(1, 13)

    rows = []
    emp_counter = 1

    for dept, cfg in DEPARTMENTS.items():
        emp_ids = [f"EMP{str(emp_counter + i).zfill(4)}" for i in range(cfg["size"])]
        emp_names = {emp_id: f"Employee {emp_id[-4:]}" for emp_id in emp_ids}
        emp_counter += cfg["size"]

        for month in months:
            w_days = _working_days_in_month(year, month)
            # Gradually worsen some departments in H2 to trigger risk indicators
            att_factor = 1.0
            ot_factor  = 1.0
            if dept in ("Dyeing", "Knitting", "Logistics"):
                if month >= 7:
                    # Absenteeism increases progressively from July
                    att_factor = max(0.80, 1.0 - (month - 6) * 0.02)
                    ot_factor  = 1.0 + (month - 6) * 0.04  # OT also creeps up
            if dept == "Maintenance":
                if month >= 8:
                    ot_factor = 1.0 + (month - 7) * 0.05   # OT spikes in H2

            for emp in emp_ids:
                for d in w_days:
                    r = random.random()
                    att_rate = cfg["att"] * att_factor

                    # OT rates adjusted by ot_factor
                    adj_ot1 = min(0.5, cfg["ot1"] * ot_factor)
                    adj_ot2 = min(0.5, cfg["ot2"] * ot_factor)
                    adj_ot3 = min(0.5, cfg["ot3"] * ot_factor)

                    if r < att_rate:
                        status = "P"
                        leave_type = None
                        leave_qty = 0.0
                    elif r < att_rate + 0.05:
                        status = "L"
                        leave_type = random.choice(LEAVE_TYPES)
                        leave_qty = random.choice([0.5, 1.0])
                    else:
                        status = "A"
                        leave_type = None
                        leave_qty = 0.0

                    # OT hours (only on present days)
                    ot1 = ot2 = ot3 = 0.0
                    if status == "P":
                        if random.random() < adj_ot1:
                            ot1 = round(random.uniform(1, 3), 1)
                        if random.random() < adj_ot2:
                            ot2 = round(random.uniform(1, 2), 1)
                        if random.random() < adj_ot3:
                            ot3 = round(random.uniform(0.5, 1.5), 1)

                    # Missing hours (present but short)
                    missing = 0.0
                    if status == "P" and random.random() < cfg["miss"]:
                        missing = round(random.uniform(0.5, 2.0), 1)

                    rows.append({
                        "id_card":        emp,
                        "employee_name":  emp_names[emp],
                        "department":     dept,
                        "date":           pd.Timestamp(d),
                        "status":         status,
                        "leave_type":     leave_type,
                        "leave_quantity": leave_qty,
                        "missing_hours":  missing,
                        "ot1":            ot1,
                        "ot2":            ot2,
                        "ot3":            ot3,
                    })

    df = pd.DataFrame(rows)
    df["month"] = df["date"].dt.month
    df["year"]  = df["date"].dt.year
    df["month_name"] = df["month"].apply(lambda m: MONTH_NAMES[m])
    df["month_year"] = df["date"].dt.to_period("M").astype(str)
    df["status_norm"] = df["status"].astype(str).str.strip().str.upper()
    df["is_leave"]   = _approved_leave_mask(df)
    df["is_absent"]  = _absent_mask(df, df["status_norm"]) & ~df["is_leave"]
    df["is_present"] = df["status_norm"].isin(PRESENT_VALUES) & ~df["is_absent"] & ~df["is_leave"]
    df["leave_quantity_approved"] = np.where(
        df["is_leave"],
        np.where(df["leave_quantity"] > 0, df["leave_quantity"], 1.0),
        0.0,
    )
    df["total_ot"]   = df["ot1"] + df["ot2"] + df["ot3"]
    return df


# ─────────────────────────────────────────────────────────────────────────────
# Data loading & normalization
# ─────────────────────────────────────────────────────────────────────────────
_cached_df: pd.DataFrame | None = None


def _resolve_column(df_cols: list, aliases: list) -> str | None:
    """Return the first alias that matches a column in df_cols (case-insensitive)."""
    lower_map = {c.lower(): c for c in df_cols}
    for a in aliases:
        if a.lower() in lower_map:
            return lower_map[a.lower()]
    return None


def load_data() -> pd.DataFrame:
    global _cached_df
    if _cached_df is not None:
        return _cached_df

    df = None
    if os.path.isfile(EXCEL_FILE_PATH):
        try:
            df = pd.read_excel(EXCEL_FILE_PATH, engine="openpyxl")
            df = _normalize_excel(df)
        except Exception as exc:
            print(f"[WARNING] Could not read Excel file: {exc}")
            df = None

    if df is None:
        print("[INFO] Using generated sample data.")
        df = generate_sample_data()

    _cached_df = df
    return df


def _normalized_leave_type(df: pd.DataFrame) -> pd.Series:
    if "leave_type" not in df.columns:
        return pd.Series("", index=df.index, dtype="object")
    return df["leave_type"].fillna("").astype(str).str.strip().str.upper()


def _absent_mask(df: pd.DataFrame, status_norm: pd.Series | None = None) -> pd.Series:
    if status_norm is None:
        status_norm = (
            df["status_norm"]
            if "status_norm" in df.columns
            else (
                df["status"].astype(str).str.strip().str.upper()
                if "status" in df.columns
                else pd.Series("", index=df.index, dtype="object")
            )
        )
    leave_type_norm = _normalized_leave_type(df)
    return status_norm.isin(ABSENT_VALUES) | leave_type_norm.isin(ABSENT_VALUES)


def _normalize_excel(raw: pd.DataFrame) -> pd.DataFrame:
    """Rename Excel columns to internal names and add derived fields."""
    rename = {}
    cols = list(raw.columns)
    for internal, aliases in COLUMN_ALIASES.items():
        found = _resolve_column(cols, aliases)
        if found and found != internal:
            rename[found] = internal
    df = raw.rename(columns=rename)

    # Ensure required columns exist
    for col in ("leave_quantity", "missing_hours", "ot1", "ot2", "ot3"):
        if col not in df.columns:
            df[col] = 0.0
    if "employee_name" not in df.columns:
        df["employee_name"] = ""

    # Date handling
    if "date" in df.columns:
        df["date"] = pd.to_datetime(df["date"], errors="coerce")
        df["month"] = df["date"].dt.month
        df["year"]  = df["date"].dt.year
    else:
        # Fallback: month/year columns
        if "month" not in df.columns:
            df["month"] = 1
        if "year" not in df.columns:
            df["year"] = datetime.now().year

    df["month_name"] = df["month"].apply(lambda m: MONTH_NAMES[int(m)] if 1 <= int(m) <= 12 else "?")
    df["month_year"] = df["year"].astype(str) + "-" + df["month"].astype(str).str.zfill(2)

    # Normalize numeric columns
    for col in ("leave_quantity", "missing_hours", "ot1", "ot2", "ot3"):
        df[col] = pd.to_numeric(df[col], errors="coerce").fillna(0)
    df["employee_name"] = df["employee_name"].fillna("").astype(str).str.strip()
    if "id_card" in df.columns:
        df.loc[df["employee_name"] == "", "employee_name"] = df["id_card"].astype(str)

    # Status flags
    if "status" in df.columns:
        df["status_norm"] = df["status"].astype(str).str.strip().str.upper()
        df["is_leave"]   = _approved_leave_mask(df)
        df["is_absent"]  = _absent_mask(df, df["status_norm"]) & ~df["is_leave"]
        df["is_present"] = df["status_norm"].isin(PRESENT_VALUES) & ~df["is_absent"] & ~df["is_leave"]
    else:
        df["status_norm"] = "P"
        df["is_leave"]   = False
        df["is_absent"]  = _absent_mask(df, df["status_norm"])
        df["is_present"] = ~df["is_absent"]

    df["leave_quantity_approved"] = np.where(
        df["is_leave"],
        np.where(df["leave_quantity"] > 0, df["leave_quantity"], 1.0),
        0.0,
    )
    df["total_ot"] = df["ot1"] + df["ot2"] + df["ot3"]
    return df


def _approved_leave_mask(df: pd.DataFrame) -> pd.Series:
    """Return a boolean mask: True for rows that should be treated as approved leave.

    A row is leave when:
      • its status is a recognised leave code (LEAVE_VALUES), OR
      • it has a non-empty leave_type that is NOT an absence marker (NA, UA, etc.), OR
      • it has a positive leave_quantity value.

    NA and UA in the leave_type column are absence markers, not leave, so they
    are explicitly excluded from leave_type-based detection.
    """
    status_norm = (
        df["status_norm"]
        if "status_norm" in df.columns
        else (
            df["status"].astype(str).str.strip().str.upper()
            if "status" in df.columns
            else pd.Series("", index=df.index, dtype="object")
        )
    )
    leave_type = (
        df["leave_type"].fillna("").astype(str).str.strip()
        if "leave_type" in df.columns
        else pd.Series("", index=df.index, dtype="object")
    )
    leave_type_upper = leave_type.str.upper()

    leave_qty = (
        pd.to_numeric(df["leave_quantity"], errors="coerce").fillna(0)
        if "leave_quantity" in df.columns
        else pd.Series(0.0, index=df.index, dtype="float64")
    )

    is_leave_status      = status_norm.isin(LEAVE_VALUES)
    has_valid_leave_type = (leave_type != "") & ~leave_type_upper.isin(ABSENT_VALUES)
    has_leave_quantity   = leave_qty > 0

    return is_leave_status | has_valid_leave_type | has_leave_quantity


def _month_order(values: list) -> list[int]:
    ordered = []
    for v in values:
        try:
            m = int(v)
        except (TypeError, ValueError):
            continue
        if 1 <= m <= 12 and m not in ordered:
            ordered.append(m)
    return sorted(ordered)


def _employee_lookup(df: pd.DataFrame) -> dict:
    if "id_card" not in df.columns:
        return {}
    sub = df[["id_card", "employee_name"]].drop_duplicates()
    return {
        str(r["id_card"]): str(r["employee_name"] or r["id_card"])
        for _, r in sub.iterrows()
    }


def _employee_name(df: pd.DataFrame, emp_id: str) -> str:
    if "employee_name" not in df.columns:
        return emp_id
    sub = df[df["id_card"] == emp_id]["employee_name"]
    if sub.empty:
        return emp_id
    val = str(sub.iloc[0]).strip()
    return val if val else emp_id


# ─────────────────────────────────────────────────────────────────────────────
# Filter helpers
# ─────────────────────────────────────────────────────────────────────────────

def parse_filters(req) -> dict:
    return {
        "departments": req.args.getlist("departments") or [],
        "months":      [int(m) for m in req.args.getlist("months") if m.isdigit()],
        "years":       [int(y) for y in req.args.getlist("years") if y.isdigit()],
        "id_cards":    req.args.getlist("id_cards") or [],
        "leave_types": req.args.getlist("leave_types") or [],
        "date_from":   req.args.get("date_from", ""),
        "date_to":     req.args.get("date_to", ""),
    }


def apply_filters(df: pd.DataFrame, f: dict) -> pd.DataFrame:
    if f.get("departments"):
        df = df[df["department"].isin(f["departments"])]
    if f.get("months"):
        df = df[df["month"].isin(f["months"])]
    if f.get("years"):
        df = df[df["year"].isin(f["years"])]
    if f.get("id_cards"):
        df = df[df["id_card"].isin(f["id_cards"])]
    if f.get("leave_types") and "leave_type" in df.columns:
        df = df[df["leave_type"].isin(f["leave_types"])]
    if f.get("date_from") and "date" in df.columns:
        df = df[df["date"] >= pd.to_datetime(f["date_from"], errors="coerce")]
    if f.get("date_to") and "date" in df.columns:
        df = df[df["date"] <= pd.to_datetime(f["date_to"], errors="coerce")]
    return df


def _to_json_safe(value):
    """Convert NumPy/Pandas scalar values into JSON-serializable Python types."""
    if isinstance(value, dict):
        return {k: _to_json_safe(v) for k, v in value.items()}
    if isinstance(value, (list, tuple, set)):
        return [_to_json_safe(v) for v in value]
    if isinstance(value, np.generic):
        return value.item()
    if isinstance(value, (pd.Timestamp, datetime, date)):
        return value.isoformat()
    return value


# ─────────────────────────────────────────────────────────────────────────────
# Metric calculation helpers
# ─────────────────────────────────────────────────────────────────────────────

def _agg_metrics(group: pd.DataFrame) -> dict:
    """Aggregate one group into a metrics dict."""
    n = len(group)
    if n == 0:
        return {}
    expected_hrs = n * 8

    present = group["is_present"].sum() if "is_present" in group else n
    absent  = group["is_absent"].sum()  if "is_absent"  in group else 0
    leave_col = "leave_quantity_approved" if "leave_quantity_approved" in group.columns else "leave_quantity"
    leave   = group[leave_col].sum()

    ot1 = group["ot1"].sum()
    ot2 = group["ot2"].sum()
    ot3 = group["ot3"].sum()
    tot_ot = ot1 + ot2 + ot3
    miss   = group["missing_hours"].sum()

    return {
        "records":       n,
        "headcount":     group["id_card"].nunique(),
        "att_pct":       round(present / n * 100, 2),
        "abs_pct":       round(absent  / n * 100, 2),
        "abs_count":     int(absent),
        "leave_pct":     round(leave   / n * 100, 2),
        "leave_days":    round(float(leave), 1),
        "missing_pct":   round(miss    / expected_hrs * 100, 2),
        "ot1_pct":       round(ot1     / expected_hrs * 100, 2),
        "ot2_pct":       round(ot2     / expected_hrs * 100, 2),
        "ot3_pct":       round(ot3     / expected_hrs * 100, 2),
        "total_ot_pct":  round(tot_ot  / expected_hrs * 100, 2),
        "ot1_hours":     round(ot1, 1),
        "ot2_hours":     round(ot2, 1),
        "ot3_hours":     round(ot3, 1),
        "total_ot_hours": round(tot_ot, 1),
        "missing_hours": round(miss, 1),
    }


def _traffic_light(metric: str, value: float) -> str:
    t = THRESHOLDS.get(metric, {})
    if metric == "att_pct":
        if value >= t.get("green", 95):  return "green"
        if value >= t.get("yellow", 85): return "yellow"
        return "red"
    else:
        if value <= t.get("yellow", 5):  return "green"
        if value <= t.get("red", 15):    return "yellow"
        return "red"


def overall_kpis(df: pd.DataFrame) -> dict:
    m = _agg_metrics(df)
    kpis = {}
    for k in ("att_pct", "abs_pct", "leave_pct", "missing_pct", "total_ot_pct"):
        val = m.get(k, 0)
        kpis[k] = {"value": val, "light": _traffic_light(k, val)}
    kpis["headcount"] = m.get("headcount", 0)
    kpis["abs_count"] = m.get("abs_count", 0)
    kpis["leave_days"] = m.get("leave_days", 0.0)
    return kpis


def by_department(df: pd.DataFrame) -> list:
    out = []
    for dept, g in df.groupby("department"):
        m = _agg_metrics(g)
        m["department"] = dept
        out.append(m)
    return sorted(out, key=lambda x: x.get("department", ""))


def by_month(df: pd.DataFrame) -> list:
    out = []
    for (yr, mo), g in df.groupby(["year", "month"]):
        m = _agg_metrics(g)
        m["year"] = int(yr)
        m["month"] = int(mo)
        m["month_name"] = MONTH_NAMES[int(mo)]
        m["month_year"] = f"{int(yr)}-{str(int(mo)).zfill(2)}"
        out.append(m)
    return sorted(out, key=lambda x: (x["year"], x["month"]))


def by_month_dept(df: pd.DataFrame) -> list:
    out = []
    for (yr, mo, dept), g in df.groupby(["year", "month", "department"]):
        m = _agg_metrics(g)
        m.update({"year": int(yr), "month": int(mo),
                  "month_name": MONTH_NAMES[int(mo)],
                  "month_year": f"{int(yr)}-{str(int(mo)).zfill(2)}",
                  "department": dept})
        out.append(m)
    return sorted(out, key=lambda x: (x["year"], x["month"], x["department"]))


def by_employee(df: pd.DataFrame) -> list:
    out = []
    for emp, g in df.groupby("id_card"):
        m = _agg_metrics(g)
        m["id_card"] = emp
        m["employee_name"] = _employee_name(df, emp)
        dept_series = g["department"].mode()
        m["department"] = dept_series.iloc[0] if not dept_series.empty else "Unknown"
        out.append(m)
    return out


def by_leave_type_dept(df: pd.DataFrame) -> list:
    sub = df[df["is_leave"] == True].copy()
    if "leave_type" not in sub.columns:
        sub["leave_type"] = "Unspecified Leave"
    else:
        sub["leave_type"] = sub["leave_type"].fillna("").astype(str).str.strip()
        sub.loc[sub["leave_type"] == "", "leave_type"] = "Unspecified Leave"
    out = []
    for (lt, dept), g in sub.groupby(["leave_type", "department"]):
        out.append({
            "leave_type": lt,
            "department": dept,
            "total_leave_days": round(g["leave_quantity_approved"].sum(), 1),
            "employees": g["id_card"].nunique(),
        })
    return out


# ─────────────────────────────────────────────────────────────────────────────
# Manpower risk analysis
# ─────────────────────────────────────────────────────────────────────────────

def _linear_trend(values: list) -> float:
    """Return slope of linear fit; positive = increasing."""
    if len(values) < 2:
        return 0.0
    x = np.arange(len(values), dtype=float)
    y = np.array(values, dtype=float)
    if np.std(y) == 0:
        return 0.0
    coeffs = np.polyfit(x, y, 1)
    return float(coeffs[0])


def manpower_risk(monthly_dept: list) -> dict:
    """Identify at-risk departments based on trends."""
    import collections

    # Group by department
    dept_months: dict = collections.defaultdict(list)
    for row in sorted(monthly_dept, key=lambda r: (r["year"], r["month"])):
        dept_months[row["department"]].append(row)

    risks = {
        "increasing_absenteeism": [],
        "increasing_ot": [],
        "low_attendance_consecutive": [],
        "excessive_leave": [],
        "high_missing_hours": [],
        "intervention_needed": [],
    }

    for dept, months in dept_months.items():
        if len(months) < 3:
            continue
        last3 = months[-3:]
        abs_slope  = _linear_trend([r["abs_pct"] for r in last3])
        ot_slope   = _linear_trend([r["total_ot_pct"] for r in last3])
        att_values = [r["att_pct"] for r in months]
        leave_avg  = np.mean([r["leave_pct"] for r in months])
        miss_avg   = np.mean([r["missing_pct"] for r in months])

        if abs_slope > 0.5:
            risks["increasing_absenteeism"].append({"department": dept, "slope": round(abs_slope, 2)})
        if ot_slope > 0.5:
            risks["increasing_ot"].append({"department": dept, "slope": round(ot_slope, 2)})

        # 2+ consecutive months with attendance < 85%
        consec = 0
        max_consec = 0
        for a in att_values:
            consec = consec + 1 if a < 85 else 0
            max_consec = max(max_consec, consec)
        if max_consec >= 2:
            risks["low_attendance_consecutive"].append({"department": dept, "months": max_consec})

        if leave_avg > 10:
            risks["excessive_leave"].append({"department": dept, "leave_pct": round(leave_avg, 2)})
        if miss_avg > 7:
            risks["high_missing_hours"].append({"department": dept, "missing_pct": round(miss_avg, 2)})

    # Departments needing intervention: appear in multiple risk categories
    from collections import Counter
    all_flagged = (
        [r["department"] for r in risks["increasing_absenteeism"]] +
        [r["department"] for r in risks["low_attendance_consecutive"]] +
        [r["department"] for r in risks["excessive_leave"]]
    )
    for dept, cnt in Counter(all_flagged).items():
        if cnt >= 2:
            risks["intervention_needed"].append({"department": dept, "risk_count": cnt})

    return risks


# ─────────────────────────────────────────────────────────────────────────────
# Exception reports
# ─────────────────────────────────────────────────────────────────────────────

def exception_reports(df: pd.DataFrame, emp_metrics: list) -> dict:
    # Employees with no leave
    all_leave = set(df[df["leave_quantity_approved"] > 0]["id_card"].unique())
    all_emp   = set(df["id_card"].unique())
    no_leave  = [
        {
            "id_card": e,
            "employee_name": _employee_name(df, e),
            "department": _emp_dept(df, e),
        }
        for e in (all_emp - all_leave)
    ]

    # Excessive OT (total_ot_pct > 25%)
    exc_ot = [
        {"id_card": e["id_card"], "department": e["department"],
         "employee_name": e.get("employee_name", e["id_card"]),
         "total_ot_pct": e["total_ot_pct"], "total_ot_hours": e["total_ot_hours"]}
        for e in emp_metrics if e.get("total_ot_pct", 0) > 25
    ]

    # Excessive missing hours (missing_pct > 10%)
    exc_miss = [
        {"id_card": e["id_card"], "department": e["department"],
         "employee_name": e.get("employee_name", e["id_card"]),
         "missing_pct": e["missing_pct"], "missing_hours": e["missing_hours"]}
        for e in emp_metrics if e.get("missing_pct", 0) > 10
    ]

    # Excessive absenteeism (abs_pct > 15%)
    exc_abs = [
        {"id_card": e["id_card"], "department": e["department"], "employee_name": e.get("employee_name", e["id_card"]), "abs_pct": e["abs_pct"]}
        for e in emp_metrics if e.get("abs_pct", 0) > 15
    ]

    # Employees with 3+ consecutive absences
    consec_abs = _find_consecutive_absences(df, min_days=3)

    return {
        "no_leave":            sorted(no_leave, key=lambda x: x["id_card"])[:50],
        "excessive_ot":        sorted(exc_ot, key=lambda x: -x["total_ot_pct"])[:50],
        "excessive_missing":   sorted(exc_miss, key=lambda x: -x["missing_pct"])[:50],
        "excessive_absenteeism": sorted(exc_abs, key=lambda x: -x["abs_pct"])[:50],
        "consecutive_absences": consec_abs[:50],
    }


def _emp_dept(df: pd.DataFrame, emp: str) -> str:
    sub = df[df["id_card"] == emp]["department"]
    return sub.mode().iloc[0] if not sub.empty else "Unknown"


def _find_consecutive_absences(df: pd.DataFrame, min_days: int = 3) -> list:
    if "date" not in df.columns or "is_absent" not in df.columns:
        return []
    results = []
    for emp, g in df.sort_values("date").groupby("id_card"):
        absent_dates = sorted(g[g["is_absent"] == True]["date"].tolist())
        if len(absent_dates) < min_days:
            continue
        streak = 1
        max_s  = 1
        for i in range(1, len(absent_dates)):
            diff = (absent_dates[i] - absent_dates[i - 1]).days
            if diff <= 2:  # allow weekends
                streak += 1
                max_s = max(max_s, streak)
            else:
                streak = 1
        if max_s >= min_days:
            dept = _emp_dept(df, emp)
            results.append({
                "id_card": emp,
                "employee_name": _employee_name(df, emp),
                "department": dept,
                "max_consecutive": max_s,
            })
    return sorted(results, key=lambda x: -x["max_consecutive"])


def _last_data_refresh_date(df: pd.DataFrame) -> str:
    if "date" in df.columns:
        dmax = pd.to_datetime(df["date"], errors="coerce").max()
        if pd.notna(dmax):
            return dmax.strftime("%Y-%m-%d")
    return datetime.now().strftime("%Y-%m-%d")


# ─────────────────────────────────────────────────────────────────────────────
# Forecasting (linear extrapolation for next 3 months)
# ─────────────────────────────────────────────────────────────────────────────

def forecast_metrics(monthly: list, horizon: int = 3) -> dict:
    if len(monthly) < 4:
        return {}

    metrics_to_forecast = ["att_pct", "abs_pct", "total_ot_pct", "missing_pct"]
    forecasts = {}

    for metric in metrics_to_forecast:
        vals = [r[metric] for r in monthly]
        x = np.arange(len(vals), dtype=float)
        coeffs = np.polyfit(x, vals, 1)
        future_x = np.arange(len(vals), len(vals) + horizon, dtype=float)
        future_y = np.polyval(coeffs, future_x)

        # Generate future month labels
        last = monthly[-1]
        future_labels = []
        mo, yr = last["month"], last["year"]
        for _ in range(horizon):
            mo += 1
            if mo > 12:
                mo = 1
                yr += 1
            future_labels.append(f"{MONTH_NAMES[mo]} {yr}")

        forecasts[metric] = {
            "historical":  [r[metric] for r in monthly],
            "labels":      [r["month_name"] + " " + str(r["year"]) for r in monthly],
            "future_vals": [round(max(0, min(100, v)), 2) for v in future_y],
            "future_labels": future_labels,
            "trend": "increasing" if coeffs[0] > 0 else "decreasing",
            "slope": round(float(coeffs[0]), 4),
        }

    return forecasts


# ─────────────────────────────────────────────────────────────────────────────
# Flask application
# ─────────────────────────────────────────────────────────────────────────────
app = Flask(__name__)


@app.route("/")
def index():
    df = load_data()
    filters_meta = {
        "departments": sorted(df["department"].unique().tolist()),
        "months":      _month_order(df["month"].unique().tolist()),
        "years":       sorted(df["year"].unique().tolist()),
        "id_cards":    sorted(df["id_card"].unique().tolist()),
        "leave_types": sorted(df["leave_type"].dropna().unique().tolist()) if "leave_type" in df.columns else [],
        "employees":   [{"id_card": k, "employee_name": v} for k, v in sorted(_employee_lookup(df).items())],
    }
    return render_template("index.html", filters_meta=filters_meta)


@app.route("/drilldown")
def drilldown():
    df = load_data()
    filters_meta = {
        "departments": sorted(df["department"].unique().tolist()),
        "months":      _month_order(df["month"].unique().tolist()),
        "years":       sorted(df["year"].unique().tolist()),
        "id_cards":    sorted(df["id_card"].unique().tolist()),
        "leave_types": sorted(df["leave_type"].dropna().unique().tolist()) if "leave_type" in df.columns else [],
        "employees":   [{"id_card": k, "employee_name": v} for k, v in sorted(_employee_lookup(df).items())],
    }
    return render_template("drilldown.html", filters_meta=filters_meta)


@app.route("/methodology")
def methodology():
    return render_template("methodology.html", thresholds=THRESHOLDS, approved_leave_types=sorted(APPROVED_LEAVE_TYPES))


@app.route("/api/filters")
def api_filters():
    df = load_data()
    return jsonify(_to_json_safe({
        "departments": sorted(df["department"].unique().tolist()),
        "months":      _month_order(df["month"].unique().tolist()),
        "years":       sorted(df["year"].unique().tolist()),
        "id_cards":    sorted(df["id_card"].unique().tolist()),
        "leave_types": sorted(df["leave_type"].dropna().unique().tolist()) if "leave_type" in df.columns else [],
        "employees":   [{"id_card": k, "employee_name": v} for k, v in sorted(_employee_lookup(df).items())],
    }))


@app.route("/api/dashboard")
def api_dashboard():
    df_raw = load_data()
    df = df_raw.copy()
    f  = parse_filters(request)
    df = apply_filters(df, f)

    if df.empty:
        return jsonify(_to_json_safe({
            "kpis": {},
            "best_dept": "N/A",
            "worst_dept": "N/A",
            "dept_metrics": [],
            "month_data": [],
            "mo_dept_data": [],
            "top_ot": [],
            "top_miss": [],
            "top_abs": [],
            "dept_rank_att": [],
            "dept_rank_abs": [],
            "dept_rank_ot": [],
            "dept_rank_miss": [],
            "leave_data": [],
            "risks": {},
            "forecasts": {},
            "exceptions": {
                "no_leave": [],
                "excessive_ot": [],
                "excessive_missing": [],
                "excessive_absenteeism": [],
                "consecutive_absences": [],
            },
            "headcount_trend": [],
            "no_data": True,
            "message": "No data available for the selected filters.",
            "last_data_refresh_date": _last_data_refresh_date(df_raw),
            "data_validation": {
                "total_records_loaded": int(len(df_raw)),
                "records_after_filters": 0,
            },
        })), 200

    kpis         = overall_kpis(df)
    dept_metrics = by_department(df)
    month_data   = by_month(df)
    mo_dept_data = by_month_dept(df)
    emp_metrics  = by_employee(df)
    leave_data   = by_leave_type_dept(df)
    risks        = manpower_risk(mo_dept_data)
    forecasts    = forecast_metrics(month_data)
    exceptions   = exception_reports(df, emp_metrics)

    # Top employee rankings
    top_ot   = sorted(emp_metrics, key=lambda x: -x.get("total_ot_hours", 0))[:10]
    top_miss = sorted(emp_metrics, key=lambda x: -x.get("missing_hours", 0))[:10]
    top_abs  = sorted(emp_metrics, key=lambda x: -x.get("abs_pct", 0))[:10]

    # Department rankings
    dept_rank_att  = sorted(dept_metrics, key=lambda x: -x.get("att_pct", 0))
    dept_rank_abs  = sorted(dept_metrics, key=lambda x: -x.get("abs_pct", 0))
    dept_rank_ot   = sorted(dept_metrics, key=lambda x: -x.get("total_ot_pct", 0))
    dept_rank_miss = sorted(dept_metrics, key=lambda x: -x.get("missing_pct", 0))

    # Best / needs attention based on composite department score
    dept_scores = sorted(
        [
            {
                "department": d["department"],
                "score": round(
                    d["att_pct"] - d["abs_pct"] * 0.5 - d["total_ot_pct"] * 0.3 - d["missing_pct"] * 0.5 - d["leave_pct"] * 0.3,
                    2,
                ),
            }
            for d in dept_metrics
        ],
        key=lambda x: x["score"],
        reverse=True,
    )
    best_dept = dept_scores[0]["department"] if dept_scores else "N/A"
    worst_dept = dept_scores[-1]["department"] if dept_scores else "N/A"
    if best_dept == worst_dept and len(dept_scores) > 1:
        worst_dept = dept_scores[-2]["department"]

    # Headcount trend (monthly)
    headcount_trend = [
        {"month_year": r["month_year"], "month_name": r["month_name"],
         "year": r["year"], "headcount": r["headcount"]}
        for r in month_data
    ]

    return jsonify(_to_json_safe({
        "kpis":              kpis,
        "best_dept":         best_dept,
        "worst_dept":        worst_dept,
        "dept_metrics":      dept_metrics,
        "month_data":        month_data,
        "mo_dept_data":      mo_dept_data,
        "top_ot":            top_ot,
        "top_miss":          top_miss,
        "top_abs":           top_abs,
        "dept_rank_att":     dept_rank_att[:5],
        "dept_rank_abs":     dept_rank_abs[:5],
        "dept_rank_ot":      dept_rank_ot[:5],
        "dept_rank_miss":    dept_rank_miss[:5],
        "leave_data":        leave_data,
        "risks":             risks,
        "forecasts":         forecasts,
        "exceptions":        exceptions,
        "headcount_trend":   headcount_trend,
        "no_data":           False,
        "message":           "",
        "last_data_refresh_date": _last_data_refresh_date(df_raw),
        "data_validation": {
            "total_records_loaded": int(len(df_raw)),
            "records_after_filters": int(len(df)),
        },
    }))


@app.route("/api/drilldown")
def api_drilldown():
    df_raw = load_data()
    df = df_raw.copy()
    f  = parse_filters(request)
    df = apply_filters(df, f)

    if df.empty:
        return jsonify(_to_json_safe({
            "emp_metrics": [],
            "dept_metrics": [],
            "mo_dept": [],
            "leave_data": [],
            "exceptions": {
                "no_leave": [],
                "excessive_ot": [],
                "excessive_missing": [],
                "excessive_absenteeism": [],
                "consecutive_absences": [],
            },
            "emp_monthly": [],
            "ot_breakdown": [],
            "recurring": {"recurring_absent": [], "recurring_missing": []},
            "no_data": True,
            "message": "No data available for the selected filters.",
            "last_data_refresh_date": _last_data_refresh_date(df_raw),
            "data_validation": {
                "total_records_loaded": int(len(df_raw)),
                "records_after_filters": 0,
            },
        })), 200

    emp_metrics  = by_employee(df)
    dept_metrics = by_department(df)
    mo_dept      = by_month_dept(df)
    leave_data   = by_leave_type_dept(df)
    exceptions   = exception_reports(df, emp_metrics)

    # Per-employee monthly details
    emp_monthly = []
    if f.get("id_cards"):
        for (emp, yr, mo), g in df.groupby(["id_card", "year", "month"]):
            m = _agg_metrics(g)
            m.update({"id_card": emp, "employee_name": _employee_name(df, emp), "year": int(yr), "month": int(mo),
                      "month_name": MONTH_NAMES[int(mo)]})
            emp_monthly.append(m)

    # OT breakdown by type and department
    ot_breakdown = []
    for dept, g in df.groupby("department"):
        expected = len(g) * 8
        ot_breakdown.append({
            "department": dept,
            "ot1_pct":   round(g["ot1"].sum() / expected * 100, 2) if expected else 0,
            "ot2_pct":   round(g["ot2"].sum() / expected * 100, 2) if expected else 0,
            "ot3_pct":   round(g["ot3"].sum() / expected * 100, 2) if expected else 0,
            "total_ot_pct": round((g["ot1"] + g["ot2"] + g["ot3"]).sum() / expected * 100, 2) if expected else 0,
        })

    # Recurring patterns per employee
    recurring = _recurring_employee_patterns(df)

    return jsonify(_to_json_safe({
        "emp_metrics":   emp_metrics,
        "dept_metrics":  dept_metrics,
        "mo_dept":       mo_dept,
        "leave_data":    leave_data,
        "exceptions":    exceptions,
        "emp_monthly":   emp_monthly,
        "ot_breakdown":  ot_breakdown,
        "recurring":     recurring,
        "no_data":       False,
        "message":       "",
        "last_data_refresh_date": _last_data_refresh_date(df_raw),
        "data_validation": {
            "total_records_loaded": int(len(df_raw)),
            "records_after_filters": int(len(df)),
        },
    }))


def _recurring_employee_patterns(df: pd.DataFrame) -> dict:
    """Detect employees with recurring absenteeism / missing hours."""
    abs_recurring  = []
    miss_recurring = []

    for emp, g in df.groupby("id_card"):
        mo_abs  = g.groupby(["year", "month"])["is_absent"].sum()
        mo_miss = g.groupby(["year", "month"])["missing_hours"].sum()
        n_months_with_abs  = (mo_abs > 2).sum()
        n_months_with_miss = (mo_miss > 4).sum()
        dept = _emp_dept(df, emp)

        if n_months_with_abs >= 3:
            abs_recurring.append({
                "id_card": emp, "employee_name": _employee_name(df, emp), "department": dept,
                "months_with_absences": int(n_months_with_abs)
            })
        if n_months_with_miss >= 3:
            miss_recurring.append({
                "id_card": emp, "employee_name": _employee_name(df, emp), "department": dept,
                "months_with_missing": int(n_months_with_miss)
            })

    return {
        "recurring_absent":  sorted(abs_recurring,  key=lambda x: -x["months_with_absences"])[:20],
        "recurring_missing": sorted(miss_recurring, key=lambda x: -x["months_with_missing"])[:20],
    }


@app.route("/api/export")
def api_export():
    df = load_data().copy()
    f = parse_filters(request)
    df = apply_filters(df, f)

    if request.args.get("selected_department_only", "").lower() in {"1", "true", "yes"} and f.get("departments"):
        df = df[df["department"] == f["departments"][0]]

    export_cols = [
        c for c in [
            "id_card", "employee_name", "department", "date", "year", "month", "status", "leave_type",
            "leave_quantity", "leave_quantity_approved", "is_present", "is_absent", "is_leave",
            "missing_hours", "ot1", "ot2", "ot3", "total_ot",
        ] if c in df.columns
    ]
    out_df = df[export_cols].copy()
    if "date" in out_df.columns:
        out_df["date"] = pd.to_datetime(out_df["date"], errors="coerce").dt.strftime("%Y-%m-%d")

    output = BytesIO()
    with pd.ExcelWriter(output, engine="openpyxl") as writer:
        out_df.to_excel(writer, index=False, sheet_name="AttendanceData")
    output.seek(0)

    stamp = datetime.now().strftime("%Y%m%d_%H%M%S")
    return send_file(
        output,
        mimetype="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        as_attachment=True,
        download_name=f"attendance_export_{stamp}.xlsx",
    )


# ─────────────────────────────────────────────────────────────────────────────
if __name__ == "__main__":
    app.run(host="0.0.0.0", port=PORT, debug=DEBUG)

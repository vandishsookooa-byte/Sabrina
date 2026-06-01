import json
import os
from dataclasses import dataclass
from typing import Dict, List

import numpy as np
import pandas as pd
from flask import Flask, render_template, request


app = Flask(__name__)

DEFAULT_EXCEL_PATH = r"C:\Users\transport\Desktop\SABRINA ATTENDANCE\Last 1 year attendance.xlsx"
EXPECTED_DAILY_HOURS = 8.0

COLUMN_ALIASES = {
    "date": ["date", "attendance date", "work date", "day"],
    "department": ["department", "dept", "business unit"],
    "employee": ["employee", "employee name", "name", "staff", "associate"],
    "leave_type": ["leave type", "leave_type", "leave"],
    "leave_qty": ["leave quantity", "leave qty", "leave_quantity", "leave days", "leave hours"],
    "missing_hours": ["missing hours", "missing_hours", "short hours"],
    "attendance": ["attendance", "status", "attendance status"],
    "ot1": ["ot1", "ot 1"],
    "ot2": ["ot2", "ot 2"],
    "ot3": ["ot3", "ot 3"],
}

ABSENT_MARKERS = {"absent", "a", "unauthorized absence"}
PRESENT_MARKERS = {"present", "p", "worked"}


@dataclass
class DashboardPayload:
    summary: Dict[str, object]
    monthly: pd.DataFrame
    rankings: pd.DataFrame
    monthly_ranked: pd.DataFrame
    top_absenteeism: pd.DataFrame
    top_missing: pd.DataFrame
    top_overtime: pd.DataFrame
    leave_utilization: pd.DataFrame
    absenteeism_ot_corr: pd.DataFrame
    chart_data: Dict[str, dict]
    filter_values: Dict[str, List[str]]
    selected_filters: Dict[str, str]


def normalize_columns(df: pd.DataFrame) -> pd.DataFrame:
    renamed = {}
    lower_map = {c.strip().lower(): c for c in df.columns}
    for canonical, aliases in COLUMN_ALIASES.items():
        for alias in aliases:
            if alias in lower_map:
                renamed[lower_map[alias]] = canonical
                break
    return df.rename(columns=renamed)


def ensure_column(df: pd.DataFrame, column: str, default):
    if column not in df.columns:
        df[column] = default


def prepare_data(df: pd.DataFrame) -> pd.DataFrame:
    df = normalize_columns(df.copy())
    ensure_column(df, "date", pd.NaT)
    ensure_column(df, "department", "Unknown")
    ensure_column(df, "employee", "Unknown")
    ensure_column(df, "leave_type", "None")
    ensure_column(df, "leave_qty", 0)
    ensure_column(df, "missing_hours", 0)
    ensure_column(df, "attendance", "")
    ensure_column(df, "ot1", 0)
    ensure_column(df, "ot2", 0)
    ensure_column(df, "ot3", 0)

    df["date"] = pd.to_datetime(df["date"], errors="coerce")
    df["department"] = df["department"].fillna("Unknown").astype(str)
    df["employee"] = df["employee"].fillna("Unknown").astype(str)
    df["leave_type"] = df["leave_type"].fillna("None").astype(str)
    df["attendance"] = df["attendance"].fillna("").astype(str).str.lower().str.strip()

    for col in ["leave_qty", "missing_hours", "ot1", "ot2", "ot3"]:
        df[col] = pd.to_numeric(df[col], errors="coerce").fillna(0.0)

    df["month"] = df["date"].dt.to_period("M").astype(str)
    df["total_ot"] = df["ot1"] + df["ot2"] + df["ot3"]
    df["is_absent"] = df["attendance"].isin(ABSENT_MARKERS) | (
        (~df["attendance"].isin(PRESENT_MARKERS)) & (df["leave_qty"] > 0)
    )
    df["is_present"] = df["attendance"].isin(PRESENT_MARKERS)
    return df


def apply_filters(df: pd.DataFrame, args) -> pd.DataFrame:
    filtered = df.copy()
    start_date = args.get("start_date")
    end_date = args.get("end_date")
    department = args.get("department")
    employee = args.get("employee")
    leave_type = args.get("leave_type")

    if start_date:
        filtered = filtered[filtered["date"] >= pd.to_datetime(start_date, errors="coerce")]
    if end_date:
        filtered = filtered[filtered["date"] <= pd.to_datetime(end_date, errors="coerce")]
    if department:
        filtered = filtered[filtered["department"] == department]
    if employee:
        filtered = filtered[filtered["employee"] == employee]
    if leave_type:
        filtered = filtered[filtered["leave_type"] == leave_type]
    return filtered


def build_monthly_metrics(df: pd.DataFrame) -> pd.DataFrame:
    if df.empty:
        return pd.DataFrame(
            columns=[
                "month",
                "department",
                "ot1_pct",
                "ot2_pct",
                "ot3_pct",
                "leave_pct",
                "missing_hours_pct",
                "attendance_pct",
                "absenteeism_pct",
                "overtime_hours",
            ]
        )

    grouped = df.groupby(["month", "department"], dropna=False).agg(
        records=("employee", "count"),
        present_count=("is_present", "sum"),
        absent_count=("is_absent", "sum"),
        leave_qty=("leave_qty", "sum"),
        missing_hours=("missing_hours", "sum"),
        ot1=("ot1", "sum"),
        ot2=("ot2", "sum"),
        ot3=("ot3", "sum"),
        total_ot=("total_ot", "sum"),
    )
    grouped = grouped.reset_index()
    total_possible_hours = np.where(grouped["records"] > 0, grouped["records"] * EXPECTED_DAILY_HOURS, np.nan)
    grouped["ot1_pct"] = (grouped["ot1"] / total_possible_hours) * 100
    grouped["ot2_pct"] = (grouped["ot2"] / total_possible_hours) * 100
    grouped["ot3_pct"] = (grouped["ot3"] / total_possible_hours) * 100
    grouped["leave_pct"] = np.where(grouped["records"] > 0, (grouped["leave_qty"] / grouped["records"]) * 100, np.nan)
    grouped["missing_hours_pct"] = (grouped["missing_hours"] / total_possible_hours) * 100
    grouped["attendance_pct"] = np.where(grouped["records"] > 0, (grouped["present_count"] / grouped["records"]) * 100, np.nan)
    grouped["absenteeism_pct"] = np.where(grouped["records"] > 0, (grouped["absent_count"] / grouped["records"]) * 100, np.nan)
    grouped["overtime_hours"] = grouped["total_ot"]

    metrics = grouped[
        [
            "month",
            "department",
            "ot1_pct",
            "ot2_pct",
            "ot3_pct",
            "leave_pct",
            "missing_hours_pct",
            "attendance_pct",
            "absenteeism_pct",
            "overtime_hours",
        ]
    ].copy()
    metrics = metrics.fillna(0.0)
    return metrics.sort_values(["month", "department"]).reset_index(drop=True)


def build_rankings(monthly: pd.DataFrame) -> pd.DataFrame:
    if monthly.empty:
        return pd.DataFrame()
    summary = monthly.groupby("department", dropna=False).agg(
        attendance_pct=("attendance_pct", "mean"),
        absenteeism_pct=("absenteeism_pct", "mean"),
        missing_hours_pct=("missing_hours_pct", "mean"),
        overtime_pct=("overtime_hours", "mean"),
    )
    summary = summary.reset_index()
    summary["rank_attendance"] = summary["attendance_pct"].rank(method="dense", ascending=False).astype(int)
    summary["rank_absenteeism"] = summary["absenteeism_pct"].rank(method="dense", ascending=True).astype(int)
    summary["rank_missing_hours"] = summary["missing_hours_pct"].rank(method="dense", ascending=True).astype(int)
    summary["rank_overtime"] = summary["overtime_pct"].rank(method="dense", ascending=False).astype(int)
    return summary.sort_values("rank_attendance").reset_index(drop=True)


def build_top_employees(df: pd.DataFrame):
    if df.empty:
        empty = pd.DataFrame(columns=["employee", "department", "value"])
        return empty, empty, empty
    by_emp = df.groupby(["employee", "department"], dropna=False).agg(
        absenteeism_days=("is_absent", "sum"),
        missing_hours=("missing_hours", "sum"),
        overtime_hours=("total_ot", "sum"),
    )
    by_emp = by_emp.reset_index()
    top_absenteeism = by_emp.sort_values("absenteeism_days", ascending=False).head(10)
    top_missing = by_emp.sort_values("missing_hours", ascending=False).head(10)
    top_overtime = by_emp.sort_values("overtime_hours", ascending=False).head(10)
    return top_absenteeism, top_missing, top_overtime


def build_leave_utilization(df: pd.DataFrame) -> pd.DataFrame:
    if df.empty:
        return pd.DataFrame(columns=["department", "leave_type", "leave_qty"])
    util = df.groupby(["department", "leave_type"], dropna=False)["leave_qty"].sum().reset_index()
    return util.sort_values(["department", "leave_qty"], ascending=[True, False]).reset_index(drop=True)


def build_absenteeism_overtime_corr(monthly: pd.DataFrame) -> pd.DataFrame:
    if monthly.empty:
        return pd.DataFrame(columns=["department", "correlation"])
    rows = []
    for dept, grp in monthly.groupby("department"):
        if len(grp) < 2:
            corr = np.nan
        else:
            corr = grp["absenteeism_pct"].corr(grp["overtime_hours"])
        rows.append({"department": dept, "correlation": corr})
    out = pd.DataFrame(rows).fillna(0.0)
    return out.sort_values("correlation", ascending=False).reset_index(drop=True)


def chart_series(monthly: pd.DataFrame, value_col: str, title: str) -> dict:
    if monthly.empty:
        return {"title": title, "labels": [], "datasets": []}
    labels = sorted(monthly["month"].dropna().unique().tolist())
    datasets = []
    for dept, grp in monthly.groupby("department"):
        data_map = {r["month"]: float(r[value_col]) for _, r in grp.iterrows()}
        datasets.append({"label": dept, "data": [round(data_map.get(m, 0.0), 2) for m in labels]})
    return {"title": title, "labels": labels, "datasets": datasets}


def build_summary(monthly: pd.DataFrame, top_missing: pd.DataFrame, top_overtime: pd.DataFrame, rankings: pd.DataFrame) -> Dict[str, object]:
    if monthly.empty:
        return {
            "overall_attendance_pct": 0,
            "overall_absenteeism_pct": 0,
            "overall_ot_pct": 0,
            "overall_leave_pct": 0,
            "overall_missing_hours_pct": 0,
            "best_department": "N/A",
            "needs_attention_department": "N/A",
            "top_missing_employees": [],
            "top_overtime_employees": [],
        }
    overall = monthly[["attendance_pct", "absenteeism_pct", "leave_pct", "missing_hours_pct", "ot1_pct", "ot2_pct", "ot3_pct"]].mean()
    best_department = rankings.sort_values(["rank_attendance", "rank_absenteeism", "rank_missing_hours"]).head(1)["department"].iloc[0]
    needs_attention = rankings.sort_values(["rank_attendance", "rank_absenteeism"], ascending=[True, False]).tail(1)["department"].iloc[0]
    return {
        "overall_attendance_pct": round(float(overall["attendance_pct"]), 2),
        "overall_absenteeism_pct": round(float(overall["absenteeism_pct"]), 2),
        "overall_ot_pct": round(float(overall["ot1_pct"] + overall["ot2_pct"] + overall["ot3_pct"]), 2),
        "overall_leave_pct": round(float(overall["leave_pct"]), 2),
        "overall_missing_hours_pct": round(float(overall["missing_hours_pct"]), 2),
        "best_department": best_department,
        "needs_attention_department": needs_attention,
        "top_missing_employees": top_missing.head(5).to_dict(orient="records"),
        "top_overtime_employees": top_overtime.head(5).to_dict(orient="records"),
    }


def build_dashboard(df: pd.DataFrame, args) -> DashboardPayload:
    filtered = apply_filters(df, args)
    monthly = build_monthly_metrics(filtered)
    rankings = build_rankings(monthly)
    top_absenteeism, top_missing, top_overtime = build_top_employees(filtered)
    leave_util = build_leave_utilization(filtered)
    corr = build_absenteeism_overtime_corr(monthly)

    charts = {
        "ot1": chart_series(monthly, "ot1_pct", "Monthly OT1 % by Department"),
        "ot2": chart_series(monthly, "ot2_pct", "Monthly OT2 % by Department"),
        "ot3": chart_series(monthly, "ot3_pct", "Monthly OT3 % by Department"),
        "leave": chart_series(monthly, "leave_pct", "Monthly Leave % by Department"),
        "missing": chart_series(monthly, "missing_hours_pct", "Monthly Missing Hours % by Department"),
        "attendance": chart_series(monthly, "attendance_pct", "Monthly Attendance % by Department"),
        "absenteeism": chart_series(monthly, "absenteeism_pct", "Monthly Absenteeism % by Department"),
        "ot_total": chart_series(monthly, "overtime_hours", "Monthly Overtime Hours by Department"),
    }
    filter_values = {
        "departments": sorted(df["department"].dropna().astype(str).unique().tolist()),
        "employees": sorted(df["employee"].dropna().astype(str).unique().tolist()),
        "leave_types": sorted(df["leave_type"].dropna().astype(str).unique().tolist()),
    }
    selected_filters = {
        "start_date": args.get("start_date", ""),
        "end_date": args.get("end_date", ""),
        "department": args.get("department", ""),
        "employee": args.get("employee", ""),
        "leave_type": args.get("leave_type", ""),
    }
    summary = build_summary(monthly, top_missing, top_overtime, rankings if not rankings.empty else pd.DataFrame({"department": ["N/A"], "rank_attendance": [1], "rank_absenteeism": [1], "rank_missing_hours": [1]}))

    monthly_ranked = monthly.sort_values(["month", "attendance_pct"], ascending=[True, False]).reset_index(drop=True)

    return DashboardPayload(
        summary=summary,
        monthly=monthly,
        rankings=rankings,
        monthly_ranked=monthly_ranked,
        top_absenteeism=top_absenteeism,
        top_missing=top_missing,
        top_overtime=top_overtime,
        leave_utilization=leave_util,
        absenteeism_ot_corr=corr,
        chart_data=charts,
        filter_values=filter_values,
        selected_filters=selected_filters,
    )


def read_attendance_file(path: str) -> pd.DataFrame:
    if not path:
        raise FileNotFoundError("No Excel file path provided.")
    if not os.path.exists(path):
        raise FileNotFoundError(f"Excel file not found: {path}")
    raw = pd.read_excel(path)
    return prepare_data(raw)


@app.route("/")
def dashboard():
    excel_path = request.args.get("file_path") or os.getenv("EXCEL_FILE_PATH", DEFAULT_EXCEL_PATH)
    error = None
    payload = None
    try:
        df = read_attendance_file(excel_path)
        payload = build_dashboard(df, request.args)
    except Exception as exc:
        error = str(exc)

    return render_template(
        "dashboard.html",
        excel_path=excel_path,
        error=error,
        payload=payload,
        chart_data_json=json.dumps(payload.chart_data if payload else {}),
    )


if __name__ == "__main__":
    port = int(os.getenv("PORT", "5000"))
    app.run(host="0.0.0.0", port=port, debug=True)

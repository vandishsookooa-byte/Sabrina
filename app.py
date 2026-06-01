import json
import os
from datetime import UTC, datetime
from typing import Dict, Iterable, List

import numpy as np
import pandas as pd
from flask import Flask, render_template, request

app = Flask(__name__)

DEFAULT_EXCEL_PATH = r"C:\Users\transport\Desktop\SABRINA ATTENDANCE\Last 1 year attendance.xlsx"
EXPECTED_DAILY_HOURS = 8.0
LEAVE_MARKERS = {"leave", "l", "half leave", "full leave"}
ABSENT_MARKERS = {"absent", "a", "unauthorized absence"}
PRESENT_MARKERS = {"present", "p", "worked"}

COLUMN_ALIASES = {
    "date": ["date", "attendance date", "work date", "day"],
    "department": ["department", "dept", "business unit"],
    "id_card": ["id card", "id", "employee id", "employee code", "emp id", "employee", "employee name", "name"],
    "leave_type": ["leave type", "leave_type", "leave"],
    "leave_qty": ["leave quantity", "leave qty", "leave_quantity", "leave days", "leave hours"],
    "missing_hours": ["missing hours", "missing_hours", "short hours"],
    "attendance": ["attendance", "status", "attendance status"],
    "ot1": ["ot1", "ot 1"],
    "ot2": ["ot2", "ot 2"],
    "ot3": ["ot3", "ot 3"],
}


def normalize_columns(df: pd.DataFrame) -> pd.DataFrame:
    lower_map = {c.strip().lower(): c for c in df.columns}
    renamed = {}
    for canonical, aliases in COLUMN_ALIASES.items():
        for alias in aliases:
            if alias in lower_map:
                renamed[lower_map[alias]] = canonical
                break
    return df.rename(columns=renamed)


def ensure_column(df: pd.DataFrame, column: str, default):
    if column not in df.columns:
        df[column] = default


def prepare_data(raw: pd.DataFrame) -> pd.DataFrame:
    df = normalize_columns(raw.copy())
    ensure_column(df, "date", pd.NaT)
    ensure_column(df, "department", "Unknown")
    ensure_column(df, "id_card", "Unknown")
    ensure_column(df, "leave_type", "None")
    ensure_column(df, "leave_qty", 0)
    ensure_column(df, "missing_hours", 0)
    ensure_column(df, "attendance", "")
    ensure_column(df, "ot1", 0)
    ensure_column(df, "ot2", 0)
    ensure_column(df, "ot3", 0)

    df["date"] = pd.to_datetime(df["date"], errors="coerce")
    df = df[df["date"].notna()].copy()
    df["department"] = df["department"].fillna("Unknown").astype(str).str.strip()
    df["id_card"] = df["id_card"].fillna("Unknown").astype(str).str.strip()
    df["leave_type"] = df["leave_type"].fillna("None").astype(str).str.strip()
    df["attendance"] = df["attendance"].fillna("").astype(str).str.lower().str.strip()

    for col in ["leave_qty", "missing_hours", "ot1", "ot2", "ot3"]:
        df[col] = pd.to_numeric(df[col], errors="coerce").fillna(0.0)

    df["month"] = df["date"].dt.to_period("M").astype(str)
    df["month_num"] = df["date"].dt.month
    df["year"] = df["date"].dt.year.astype(int)
    df["month_label"] = df["date"].dt.strftime("%b")
    df["total_ot"] = df["ot1"] + df["ot2"] + df["ot3"]

    df["is_leave"] = (df["leave_qty"] > 0) | (df["attendance"].isin(LEAVE_MARKERS))
    df["is_absent"] = df["attendance"].isin(ABSENT_MARKERS)
    df["is_present"] = df["attendance"].isin(PRESENT_MARKERS) | (~df["is_absent"] & ~df["is_leave"])
    df["man_hours"] = EXPECTED_DAILY_HOURS
    return df


def _get_multi(args, key: str) -> List[str]:
    values = [v.strip() for v in args.getlist(key) if v.strip()]
    if values:
        return values
    csv = args.get(key, "").strip()
    if csv:
        return [v.strip() for v in csv.split(",") if v.strip()]
    return []


def apply_filters(df: pd.DataFrame, args) -> pd.DataFrame:
    filtered = df.copy()
    departments = _get_multi(args, "department")
    id_cards = _get_multi(args, "id_card")
    leave_types = _get_multi(args, "leave_type")
    months = [int(v) for v in _get_multi(args, "month") if v.isdigit()]
    years = [int(v) for v in _get_multi(args, "year") if v.isdigit()]

    start_date = args.get("start_date", "").strip()
    end_date = args.get("end_date", "").strip()
    if start_date:
        filtered = filtered[filtered["date"] >= pd.to_datetime(start_date, errors="coerce")]
    if end_date:
        filtered = filtered[filtered["date"] <= pd.to_datetime(end_date, errors="coerce")]
    if departments:
        filtered = filtered[filtered["department"].isin(departments)]
    if id_cards:
        filtered = filtered[filtered["id_card"].isin(id_cards)]
    if leave_types:
        filtered = filtered[filtered["leave_type"].isin(leave_types)]
    if months:
        filtered = filtered[filtered["month_num"].isin(months)]
    if years:
        filtered = filtered[filtered["year"].isin(years)]
    return filtered


def pct(numerator, denominator) -> pd.Series:
    return np.where(denominator > 0, (numerator / denominator) * 100, 0.0)


def build_monthly_department_metrics(df: pd.DataFrame) -> pd.DataFrame:
    columns = [
        "month",
        "year",
        "month_num",
        "department",
        "records",
        "headcount",
        "attendance_pct",
        "absenteeism_pct",
        "leave_pct",
        "missing_hours_pct",
        "ot_pct",
        "ot1_pct",
        "ot2_pct",
        "ot3_pct",
        "ot1",
        "ot2",
        "ot3",
        "total_ot",
        "missing_hours",
    ]
    if df.empty:
        return pd.DataFrame(columns=columns)

    grouped = (
        df.groupby(["month", "year", "month_num", "department"], dropna=False)
        .agg(
            records=("id_card", "count"),
            headcount=("id_card", "nunique"),
            present_count=("is_present", "sum"),
            absent_count=("is_absent", "sum"),
            leave_count=("is_leave", "sum"),
            leave_qty=("leave_qty", "sum"),
            missing_hours=("missing_hours", "sum"),
            ot1=("ot1", "sum"),
            ot2=("ot2", "sum"),
            ot3=("ot3", "sum"),
            total_ot=("total_ot", "sum"),
            total_hours=("man_hours", "sum"),
        )
        .reset_index()
    )
    grouped["attendance_pct"] = pct(grouped["present_count"], grouped["records"])
    grouped["absenteeism_pct"] = pct(grouped["absent_count"], grouped["records"])
    grouped["leave_pct"] = pct(grouped["leave_count"], grouped["records"])
    grouped["missing_hours_pct"] = pct(grouped["missing_hours"], grouped["total_hours"])
    grouped["ot_pct"] = pct(grouped["total_ot"], grouped["total_hours"])
    grouped["ot1_pct"] = pct(grouped["ot1"], grouped["total_hours"])
    grouped["ot2_pct"] = pct(grouped["ot2"], grouped["total_hours"])
    grouped["ot3_pct"] = pct(grouped["ot3"], grouped["total_hours"])
    grouped = grouped[columns]
    return grouped.sort_values(["year", "month_num", "department"]).reset_index(drop=True)


def build_department_scorecard(monthly: pd.DataFrame) -> pd.DataFrame:
    if monthly.empty:
        return pd.DataFrame()
    score = (
        monthly.groupby("department", dropna=False)
        .agg(
            attendance_pct=("attendance_pct", "mean"),
            absenteeism_pct=("absenteeism_pct", "mean"),
            leave_pct=("leave_pct", "mean"),
            missing_hours_pct=("missing_hours_pct", "mean"),
            ot_pct=("ot_pct", "mean"),
        )
        .reset_index()
    )
    score["r_att"] = score["attendance_pct"].rank(ascending=False, method="dense")
    score["r_abs"] = score["absenteeism_pct"].rank(ascending=True, method="dense")
    score["r_leave"] = score["leave_pct"].rank(ascending=True, method="dense")
    score["r_miss"] = score["missing_hours_pct"].rank(ascending=True, method="dense")
    score["r_ot"] = score["ot_pct"].rank(ascending=True, method="dense")
    score["overall_rank"] = (score[["r_att", "r_abs", "r_leave", "r_miss", "r_ot"]].mean(axis=1)).rank(ascending=True, method="dense").astype(int)
    return score.sort_values("overall_rank").reset_index(drop=True)


def build_monthly_summary(monthly: pd.DataFrame) -> pd.DataFrame:
    if monthly.empty:
        return pd.DataFrame()
    out = (
        monthly.groupby(["month", "year", "month_num"], dropna=False)
        .agg(
            attendance_pct=("attendance_pct", "mean"),
            absenteeism_pct=("absenteeism_pct", "mean"),
            leave_pct=("leave_pct", "mean"),
            ot_pct=("ot_pct", "mean"),
            missing_hours_pct=("missing_hours_pct", "mean"),
            total_employees=("headcount", "sum"),
        )
        .reset_index()
        .sort_values(["year", "month_num"])
    )
    return out.reset_index(drop=True)


def top_departments(scorecard: pd.DataFrame, metric: str, n=5) -> pd.DataFrame:
    if scorecard.empty:
        return pd.DataFrame(columns=["department", metric])
    return scorecard.sort_values(metric, ascending=False)[["department", metric]].head(n).reset_index(drop=True)


def build_top_employees(df: pd.DataFrame) -> Dict[str, pd.DataFrame]:
    if df.empty:
        empty = pd.DataFrame(columns=["id_card", "department", "value"])
        return {"ot": empty, "missing": empty, "absenteeism": empty}
    emp = (
        df.groupby(["id_card", "department"], dropna=False)
        .agg(
            overtime_hours=("total_ot", "sum"),
            missing_hours=("missing_hours", "sum"),
            absenteeism_days=("is_absent", "sum"),
            leave_days=("is_leave", "sum"),
        )
        .reset_index()
    )
    return {
        "ot": emp.sort_values("overtime_hours", ascending=False).head(10).reset_index(drop=True),
        "missing": emp.sort_values("missing_hours", ascending=False).head(10).reset_index(drop=True),
        "absenteeism": emp.sort_values("absenteeism_days", ascending=False).head(10).reset_index(drop=True),
        "base": emp,
    }


def build_leave_utilization(df: pd.DataFrame) -> pd.DataFrame:
    if df.empty:
        return pd.DataFrame(columns=["department", "leave_type", "leave_qty"])
    return (
        df.groupby(["department", "leave_type"], dropna=False)["leave_qty"]
        .sum()
        .reset_index()
        .sort_values(["department", "leave_qty"], ascending=[True, False])
        .reset_index(drop=True)
    )


def build_correlation(monthly: pd.DataFrame) -> pd.DataFrame:
    if monthly.empty:
        return pd.DataFrame(columns=["department", "correlation", "avg_absenteeism_pct", "avg_ot_pct"])
    rows = []
    for dept, grp in monthly.groupby("department"):
        with np.errstate(invalid="ignore", divide="ignore"):
            corr = float(grp["absenteeism_pct"].corr(grp["ot_pct"])) if len(grp) > 1 else 0.0
        rows.append(
            {
                "department": dept,
                "correlation": 0.0 if pd.isna(corr) else corr,
                "avg_absenteeism_pct": float(grp["absenteeism_pct"].mean()),
                "avg_ot_pct": float(grp["ot_pct"].mean()),
            }
        )
    return pd.DataFrame(rows).sort_values("correlation", ascending=False).reset_index(drop=True)


def month_order(monthly: pd.DataFrame) -> List[str]:
    if monthly.empty:
        return []
    seq = monthly[["month", "year", "month_num"]].drop_duplicates().sort_values(["year", "month_num"])
    return seq["month"].tolist()


def trend_chart(monthly: pd.DataFrame, value_col: str, title: str) -> Dict[str, object]:
    if monthly.empty:
        return {"title": title, "type": "line", "labels": [], "datasets": []}
    labels = month_order(monthly)
    data = []
    for dept, grp in monthly.groupby("department"):
        by_month = dict(zip(grp["month"], grp[value_col]))
        data.append({"label": dept, "data": [round(float(by_month.get(m, 0.0)), 2) for m in labels]})
    return {"title": title, "type": "line", "labels": labels, "datasets": data}


def department_bar_chart(scorecard: pd.DataFrame, metrics: Iterable[str], title: str) -> Dict[str, object]:
    if scorecard.empty:
        return {"title": title, "type": "bar", "labels": [], "datasets": []}
    labels = scorecard["department"].tolist()
    palettes = ["#4f46e5", "#06b6d4", "#22c55e", "#f59e0b", "#ef4444", "#a855f7", "#14b8a6"]
    datasets = []
    for idx, metric in enumerate(metrics):
        datasets.append({"label": metric.replace("_pct", " %").upper(), "data": [round(float(v), 2) for v in scorecard[metric]], "backgroundColor": palettes[idx % len(palettes)]})
    return {"title": title, "type": "bar", "labels": labels, "datasets": datasets}


def traffic_light(value: float, metric: str) -> str:
    metric = metric.lower()
    if metric == "attendance_pct":
        return "green" if value >= 95 else "yellow" if value >= 90 else "red"
    return "green" if value <= 3 else "yellow" if value <= 6 else "red"


def build_ot_monthly_table(monthly: pd.DataFrame) -> pd.DataFrame:
    if monthly.empty:
        return pd.DataFrame(columns=["department", "ot1_pct", "ot2_pct", "ot3_pct", "total_ot_pct"])
    out = (
        monthly.groupby("department", dropna=False)
        .agg(ot1_pct=("ot1_pct", "mean"), ot2_pct=("ot2_pct", "mean"), ot3_pct=("ot3_pct", "mean"), total_ot_pct=("ot_pct", "mean"))
        .reset_index()
        .sort_values("total_ot_pct", ascending=False)
    )
    return out.reset_index(drop=True)


def build_matrix(monthly: pd.DataFrame, value_col: str) -> Dict[str, object]:
    if monthly.empty:
        return {"months": [], "rows": []}
    labels = month_order(monthly)
    pivot = monthly.pivot_table(index="department", columns="month", values=value_col, aggfunc="mean", fill_value=0.0)
    pivot = pivot.reindex(columns=labels, fill_value=0.0)
    rows = [{"department": dept, "values": [round(float(v), 2) for v in vals]} for dept, vals in pivot.iterrows()]
    return {"months": labels, "rows": rows}


def build_exception_reports(df: pd.DataFrame, top_emp_base: pd.DataFrame) -> Dict[str, pd.DataFrame]:
    if df.empty or top_emp_base.empty:
        empty = pd.DataFrame(columns=["id_card", "department", "value"])
        return {"no_leave": empty, "excessive_ot": empty, "excessive_missing": empty, "excessive_absent": empty}

    ot_th = top_emp_base["overtime_hours"].quantile(0.9)
    missing_th = top_emp_base["missing_hours"].quantile(0.9)
    abs_th = top_emp_base["absenteeism_days"].quantile(0.9)
    no_leave = top_emp_base[top_emp_base["leave_days"] <= 0][["id_card", "department", "leave_days"]].head(20)
    excessive_ot = top_emp_base[top_emp_base["overtime_hours"] >= ot_th][["id_card", "department", "overtime_hours"]].sort_values("overtime_hours", ascending=False).head(20)
    excessive_missing = top_emp_base[top_emp_base["missing_hours"] >= missing_th][["id_card", "department", "missing_hours"]].sort_values("missing_hours", ascending=False).head(20)
    excessive_absent = top_emp_base[top_emp_base["absenteeism_days"] >= abs_th][["id_card", "department", "absenteeism_days"]].sort_values("absenteeism_days", ascending=False).head(20)
    return {"no_leave": no_leave, "excessive_ot": excessive_ot, "excessive_missing": excessive_missing, "excessive_absent": excessive_absent}


def consecutive_absence_report(df: pd.DataFrame) -> pd.DataFrame:
    if df.empty:
        return pd.DataFrame(columns=["id_card", "department", "max_consecutive_absences"])
    rows = []
    for (id_card, dept), grp in df.sort_values("date").groupby(["id_card", "department"]):
        streak = 0
        max_streak = 0
        for is_absent in grp["is_absent"].tolist():
            streak = streak + 1 if is_absent else 0
            max_streak = max(max_streak, streak)
        if max_streak >= 3:
            rows.append({"id_card": id_card, "department": dept, "max_consecutive_absences": max_streak})
    if not rows:
        return pd.DataFrame(columns=["id_card", "department", "max_consecutive_absences"])
    return pd.DataFrame(rows).sort_values("max_consecutive_absences", ascending=False).reset_index(drop=True)


def trend_slope(grp: pd.DataFrame, value_col: str) -> float:
    if len(grp) < 2:
        return 0.0
    x = np.arange(len(grp))
    y = grp.sort_values(["year", "month_num"])[value_col].to_numpy(dtype=float)
    slope = float(np.polyfit(x, y, 1)[0])
    return slope


def department_risks(monthly: pd.DataFrame) -> Dict[str, List[str]]:
    risks = {
        "increasing_absenteeism": [],
        "increasing_overtime": [],
        "recurring_missing_hours": [],
        "low_attendance_2_months": [],
        "excessive_leave": [],
        "needs_intervention": [],
    }
    for dept, grp in monthly.groupby("department"):
        g = grp.sort_values(["year", "month_num"])
        if trend_slope(g, "absenteeism_pct") > 0.3:
            risks["increasing_absenteeism"].append(dept)
        if trend_slope(g, "ot_pct") > 0.3:
            risks["increasing_overtime"].append(dept)
        if (g["missing_hours_pct"] > 4).sum() >= 2:
            risks["recurring_missing_hours"].append(dept)
        low_flags = (g["attendance_pct"] < 90).astype(int).tolist()
        if any(low_flags[i] == 1 and low_flags[i - 1] == 1 for i in range(1, len(low_flags))):
            risks["low_attendance_2_months"].append(dept)
        if g["leave_pct"].mean() > 8:
            risks["excessive_leave"].append(dept)
    intervention = set(sum([v for k, v in risks.items() if k != "needs_intervention"], []))
    risks["needs_intervention"] = sorted(intervention)
    return risks


def monthly_forecast(monthly_summary: pd.DataFrame, value_col: str, periods=3) -> List[Dict[str, object]]:
    if len(monthly_summary) < 2:
        return []
    series = monthly_summary.sort_values(["year", "month_num"]).reset_index(drop=True)
    x = np.arange(len(series))
    y = series[value_col].astype(float).to_numpy()
    slope, intercept = np.polyfit(x, y, 1)
    last_month = pd.Period(series.iloc[-1]["month"], freq="M")
    out = []
    for i in range(1, periods + 1):
        future_month = (last_month + i).strftime("%Y-%m")
        pred = float(max(0.0, intercept + slope * (len(series) + i - 1)))
        out.append({"month": future_month, "forecast": round(pred, 2)})
    return out


def build_headcount_analysis(df: pd.DataFrame) -> Dict[str, pd.DataFrame]:
    if df.empty:
        empty = pd.DataFrame()
        return {"monthly": empty, "joiners": empty, "leavers": empty, "department": empty}
    monthly = (
        df.groupby("month")["id_card"].nunique().reset_index(name="headcount").sort_values("month")
    )
    month_people = df.groupby("month")["id_card"].apply(set).sort_index()
    join_rows, leave_rows = [], []
    prev = set()
    for month, people in month_people.items():
        joiners = people - prev
        leavers = prev - people
        join_rows.append({"month": month, "joiners": len(joiners)})
        leave_rows.append({"month": month, "leavers": len(leavers)})
        prev = people
    dept = df.groupby("department")["id_card"].nunique().reset_index(name="headcount").sort_values("headcount", ascending=False)
    return {"monthly": monthly, "joiners": pd.DataFrame(join_rows), "leavers": pd.DataFrame(leave_rows), "department": dept}


def kpi_summary(monthly: pd.DataFrame, scorecard: pd.DataFrame, filtered_df: pd.DataFrame) -> Dict[str, object]:
    if monthly.empty:
        return {
            "attendance_pct": 0,
            "absenteeism_pct": 0,
            "leave_pct": 0,
            "missing_hours_pct": 0,
            "ot_pct": 0,
            "total_headcount": 0,
            "best_department": "N/A",
            "attention_department": "N/A",
            "traffic": {},
        }
    overall = monthly[["attendance_pct", "absenteeism_pct", "leave_pct", "missing_hours_pct", "ot_pct"]].mean()
    best_department = scorecard.head(1)["department"].iloc[0] if not scorecard.empty else "N/A"
    attention_department = scorecard.sort_values("overall_rank", ascending=False).head(1)["department"].iloc[0] if not scorecard.empty else "N/A"
    result = {
        "attendance_pct": round(float(overall["attendance_pct"]), 2),
        "absenteeism_pct": round(float(overall["absenteeism_pct"]), 2),
        "leave_pct": round(float(overall["leave_pct"]), 2),
        "missing_hours_pct": round(float(overall["missing_hours_pct"]), 2),
        "ot_pct": round(float(overall["ot_pct"]), 2),
        "total_headcount": int(filtered_df["id_card"].nunique()),
        "best_department": best_department,
        "attention_department": attention_department,
    }
    result["traffic"] = {k: traffic_light(result[k], k) for k in ["attendance_pct", "absenteeism_pct", "leave_pct", "missing_hours_pct", "ot_pct"]}
    return result


def filter_values(df: pd.DataFrame) -> Dict[str, List[str]]:
    return {
        "departments": sorted(df["department"].dropna().unique().tolist()),
        "id_cards": sorted(df["id_card"].dropna().unique().tolist()),
        "leave_types": sorted(df["leave_type"].dropna().unique().tolist()),
        "months": sorted(df["month_num"].dropna().astype(int).unique().tolist()),
        "years": sorted(df["year"].dropna().astype(int).unique().tolist()),
    }


def selected_filters(args) -> Dict[str, object]:
    return {
        "department": _get_multi(args, "department"),
        "id_card": _get_multi(args, "id_card"),
        "leave_type": _get_multi(args, "leave_type"),
        "month": [int(v) for v in _get_multi(args, "month") if v.isdigit()],
        "year": [int(v) for v in _get_multi(args, "year") if v.isdigit()],
        "start_date": args.get("start_date", ""),
        "end_date": args.get("end_date", ""),
    }


def insights(kpi: Dict[str, object], risks: Dict[str, List[str]]) -> List[str]:
    items = []
    if kpi["attendance_pct"] < 90:
        items.append("Attendance is below 90%; implement targeted attendance recovery plans.")
    if kpi["ot_pct"] > 8:
        items.append("Overtime utilization is elevated; rebalance staffing and shift allocation.")
    if risks["needs_intervention"]:
        items.append(f"Departments requiring immediate intervention: {', '.join(risks['needs_intervention'][:5])}.")
    if not items:
        items.append("Core KPIs are stable; continue monitoring trend variance monthly.")
    return items


def build_payload(df: pd.DataFrame, args) -> Dict[str, object]:
    filtered = apply_filters(df, args)
    monthly = build_monthly_department_metrics(filtered)
    scorecard = build_department_scorecard(monthly)
    monthly_summary = build_monthly_summary(monthly)
    top_emp = build_top_employees(filtered)
    leave_util = build_leave_utilization(filtered)
    corr = build_correlation(monthly)
    ot_table = build_ot_monthly_table(monthly)
    exceptions = build_exception_reports(filtered, top_emp["base"] if "base" in top_emp else pd.DataFrame())
    consecutive_abs = consecutive_absence_report(filtered)
    risks = department_risks(monthly) if not monthly.empty else {k: [] for k in ["increasing_absenteeism", "increasing_overtime", "recurring_missing_hours", "low_attendance_2_months", "excessive_leave", "needs_intervention"]}
    headcount = build_headcount_analysis(filtered)
    kpi = kpi_summary(monthly, scorecard, filtered)

    forecasts = {
        "absenteeism": monthly_forecast(monthly_summary, "absenteeism_pct"),
        "overtime": monthly_forecast(monthly_summary, "ot_pct"),
        "leave": monthly_forecast(monthly_summary, "leave_pct"),
    }
    shortages = []
    for idx in range(min(len(forecasts["absenteeism"]), len(forecasts["overtime"]))):
        month = forecasts["absenteeism"][idx]["month"]
        risk_index = forecasts["absenteeism"][idx]["forecast"] + forecasts["overtime"][idx]["forecast"]
        shortages.append({"month": month, "risk_index": round(risk_index, 2)})

    charts = {
        "dept_kpis": department_bar_chart(scorecard, ["attendance_pct", "absenteeism_pct", "leave_pct", "missing_hours_pct", "ot_pct"], "Department KPI Comparison"),
        "attendance_trend": trend_chart(monthly, "attendance_pct", "Monthly Attendance Trend by Department"),
        "absenteeism_trend": trend_chart(monthly, "absenteeism_pct", "Monthly Absenteeism Trend by Department"),
        "ot_trend": trend_chart(monthly, "ot_pct", "Monthly Overtime Trend by Department"),
        "ot1_trend": trend_chart(monthly, "ot1_pct", "Monthly OT1 Trend by Department"),
        "ot2_trend": trend_chart(monthly, "ot2_pct", "Monthly OT2 Trend by Department"),
        "ot3_trend": trend_chart(monthly, "ot3_pct", "Monthly OT3 Trend by Department"),
    }

    top_dept = {
        "attendance": top_departments(scorecard, "attendance_pct"),
        "absenteeism": top_departments(scorecard, "absenteeism_pct"),
        "overtime": top_departments(scorecard, "ot_pct"),
        "missing": top_departments(scorecard, "missing_hours_pct"),
        "ot1": top_departments(ot_table.rename(columns={"total_ot_pct": "ot_pct"}), "ot1_pct"),
        "ot2": top_departments(ot_table.rename(columns={"total_ot_pct": "ot_pct"}), "ot2_pct"),
        "ot3": top_departments(ot_table.rename(columns={"total_ot_pct": "ot_pct"}), "ot3_pct"),
    }

    return {
        "kpi": kpi,
        "summary_cards": kpi,
        "monthly_summary": monthly_summary,
        "department_scorecard": scorecard,
        "top_departments": top_dept,
        "top_employees": {"ot": top_emp["ot"], "missing": top_emp["missing"], "absenteeism": top_emp["absenteeism"]},
        "correlation": corr,
        "ot_table": ot_table,
        "matrices": {
            "ot": build_matrix(monthly, "ot_pct"),
            "attendance": build_matrix(monthly, "attendance_pct"),
            "absenteeism": build_matrix(monthly, "absenteeism_pct"),
            "missing": build_matrix(monthly, "missing_hours_pct"),
            "overtime": build_matrix(monthly, "ot_pct"),
        },
        "leave_utilization": leave_util,
        "exceptions": exceptions,
        "consecutive_absences": consecutive_abs,
        "risks": risks,
        "forecasts": forecasts,
        "shortages": shortages,
        "headcount": headcount,
        "insights": insights(kpi, risks),
        "charts": charts,
        "monthly_department": monthly,
        "filters": filter_values(df),
        "selected": selected_filters(args),
        "generated_at": datetime.now(UTC).strftime("%Y-%m-%d %H:%M UTC"),
    }


def read_attendance_file(path: str) -> pd.DataFrame:
    if not path:
        raise FileNotFoundError("No Excel file path provided.")
    if not os.path.exists(path):
        raise FileNotFoundError(f"Excel file not found: {path}")
    return prepare_data(pd.read_excel(path))


def get_payload(args):
    excel_path = os.getenv("EXCEL_FILE_PATH", DEFAULT_EXCEL_PATH)
    df = read_attendance_file(excel_path)
    return build_payload(df, args)


@app.route("/")
def dashboard():
    error = None
    payload = None
    try:
        payload = get_payload(request.args)
    except Exception as exc:
        error = str(exc)
    return render_template("dashboard.html", payload=payload, error=error, chart_data_json=json.dumps(payload["charts"] if payload else {}))


@app.route("/drilldown")
def drilldown():
    error = None
    payload = None
    detail = pd.DataFrame()
    try:
        payload = get_payload(request.args)
        detail = payload["monthly_department"].copy()
    except Exception as exc:
        error = str(exc)
    return render_template(
        "drilldown.html",
        payload=payload,
        detail=detail,
        error=error,
        chart_data_json=json.dumps(payload["charts"] if payload else {}),
    )


if __name__ == "__main__":
    port = int(os.getenv("PORT", "5000"))
    debug_mode = os.getenv("FLASK_DEBUG", "0") == "1"
    app.run(host="0.0.0.0", port=port, debug=debug_mode)

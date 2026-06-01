from __future__ import annotations

import io
from dataclasses import dataclass
from typing import Any

import pandas as pd
import plotly.express as px
from flask import Flask, render_template, request
from markupsafe import Markup, escape


app = Flask(__name__)


@dataclass
class ColumnMap:
    date: str | None
    department: str | None
    employee: str | None
    leave_type: str | None
    leave_qty: str | None
    missing_hours: str | None
    ot1: str | None
    ot2: str | None
    ot3: str | None
    attendance: str | None
    absent: str | None


def _normalize(name: str) -> str:
    return "".join(ch.lower() for ch in str(name) if ch.isalnum())


def _find_col(df: pd.DataFrame, options: list[str]) -> str | None:
    normalized = {_normalize(col): col for col in df.columns}
    for option in options:
        if option in normalized:
            return normalized[option]
    for key, col in normalized.items():
        if any(option in key for option in options):
            return col
    return None


def map_columns(df: pd.DataFrame) -> ColumnMap:
    return ColumnMap(
        date=_find_col(df, ["date", "attendancedate", "workdate"]),
        department=_find_col(df, ["department", "dept"]),
        employee=_find_col(df, ["employee", "employeename", "staff", "name"]),
        leave_type=_find_col(df, ["leavetype", "leave"]),
        leave_qty=_find_col(df, ["leavequantity", "leaveqty", "leavehours", "leavedays"]),
        missing_hours=_find_col(df, ["missinghours", "shorthours"]),
        ot1=_find_col(df, ["ot1"]),
        ot2=_find_col(df, ["ot2"]),
        ot3=_find_col(df, ["ot3"]),
        attendance=_find_col(df, ["attendance", "status"]),
        absent=_find_col(df, ["absent", "absenteeism"]),
    )


def _numeric(df: pd.DataFrame, col: str | None) -> pd.Series:
    if not col or col not in df.columns:
        return pd.Series(0.0, index=df.index)
    return pd.to_numeric(df[col], errors="coerce").fillna(0.0)


def _present_absent_flags(df: pd.DataFrame, colmap: ColumnMap) -> tuple[pd.Series, pd.Series]:
    present = pd.Series(1.0, index=df.index)
    absent = pd.Series(0.0, index=df.index)

    if colmap.attendance and colmap.attendance in df.columns:
        raw = df[colmap.attendance]
        numeric = pd.to_numeric(raw, errors="coerce")
        if numeric.notna().any():
            present = numeric.fillna(0).clip(lower=0)
        else:
            text = raw.astype(str).str.lower().str.strip()
            present = text.isin(["present", "p", "true", "yes", "attended"]).astype(float)
            absent = text.isin(["absent", "a", "false", "no"]).astype(float)

    if colmap.absent and colmap.absent in df.columns:
        absent_numeric = pd.to_numeric(df[colmap.absent], errors="coerce")
        if absent_numeric.notna().any():
            absent = absent_numeric.fillna(0).clip(lower=0)
        else:
            absent_text = df[colmap.absent].astype(str).str.lower().str.strip()
            absent = absent_text.isin(["absent", "a", "true", "yes"]).astype(float)

    present = present.clip(lower=0)
    absent = absent.clip(lower=0)
    return present, absent


def prepare_dataframe(df: pd.DataFrame) -> tuple[pd.DataFrame, ColumnMap]:
    df = df.copy()
    colmap = map_columns(df)

    if colmap.date and colmap.date in df.columns:
        df[colmap.date] = pd.to_datetime(df[colmap.date], errors="coerce")
    else:
        raise ValueError("Could not locate a date column in the source file.")

    if not colmap.department:
        raise ValueError("Could not locate a department column in the source file.")

    if not colmap.employee:
        df["_employee"] = "Unknown"
        colmap.employee = "_employee"

    df["Month"] = df[colmap.date].dt.to_period("M").astype(str)
    df["OT1"] = _numeric(df, colmap.ot1)
    df["OT2"] = _numeric(df, colmap.ot2)
    df["OT3"] = _numeric(df, colmap.ot3)
    df["LeaveQty"] = _numeric(df, colmap.leave_qty)
    df["MissingHours"] = _numeric(df, colmap.missing_hours)
    df["Present"], df["Absent"] = _present_absent_flags(df, colmap)
    df["TotalOT"] = df["OT1"] + df["OT2"] + df["OT3"]
    df["TotalEvents"] = (df["Present"] + df["Absent"] + df["LeaveQty"]).replace(0, 1)
    df["OvertimeBase"] = (df["TotalOT"] + df["MissingHours"] + df["Present"]).replace(0, 1)

    df = df.dropna(subset=[colmap.date])
    return df, colmap


def apply_filters(df: pd.DataFrame, colmap: ColumnMap, filters: dict[str, Any]) -> pd.DataFrame:
    filtered = df.copy()
    if filters.get("start_date"):
        filtered = filtered[filtered[colmap.date] >= pd.to_datetime(filters["start_date"])]
    if filters.get("end_date"):
        filtered = filtered[filtered[colmap.date] <= pd.to_datetime(filters["end_date"])]
    if filters.get("department"):
        filtered = filtered[filtered[colmap.department] == filters["department"]]
    if filters.get("employee"):
        filtered = filtered[filtered[colmap.employee] == filters["employee"]]
    if filters.get("leave_type") and colmap.leave_type and colmap.leave_type in filtered.columns:
        filtered = filtered[filtered[colmap.leave_type] == filters["leave_type"]]
    return filtered


def calculate_metrics(df: pd.DataFrame, colmap: ColumnMap) -> dict[str, Any]:
    monthly = (
        df.groupby(["Month", colmap.department], as_index=False)
        .agg(
            OT1=("OT1", "sum"),
            OT2=("OT2", "sum"),
            OT3=("OT3", "sum"),
            LeaveQty=("LeaveQty", "sum"),
            MissingHours=("MissingHours", "sum"),
            Present=("Present", "sum"),
            Absent=("Absent", "sum"),
            TotalOT=("TotalOT", "sum"),
            TotalEvents=("TotalEvents", "sum"),
            OvertimeBase=("OvertimeBase", "sum"),
        )
        .rename(columns={colmap.department: "Department"})
    )

    monthly["OT1Pct"] = (monthly["OT1"] / monthly["OvertimeBase"] * 100).round(2)
    monthly["OT2Pct"] = (monthly["OT2"] / monthly["OvertimeBase"] * 100).round(2)
    monthly["OT3Pct"] = (monthly["OT3"] / monthly["OvertimeBase"] * 100).round(2)
    monthly["LeavePct"] = (monthly["LeaveQty"] / monthly["TotalEvents"] * 100).round(2)
    monthly["MissingPct"] = (monthly["MissingHours"] / monthly["OvertimeBase"] * 100).round(2)
    monthly["AttendancePct"] = (monthly["Present"] / monthly["TotalEvents"] * 100).round(2)
    monthly["AbsenteeismPct"] = (monthly["Absent"] / monthly["TotalEvents"] * 100).round(2)

    dept_summary = (
        monthly.groupby("Department", as_index=False)
        .agg(
            OT1Pct=("OT1Pct", "mean"),
            OT2Pct=("OT2Pct", "mean"),
            OT3Pct=("OT3Pct", "mean"),
            LeavePct=("LeavePct", "mean"),
            MissingPct=("MissingPct", "mean"),
            AttendancePct=("AttendancePct", "mean"),
            AbsenteeismPct=("AbsenteeismPct", "mean"),
            TotalOTHours=("TotalOT", "sum"),
        )
    ).round(2)

    by_emp = (
        df.groupby(colmap.employee, as_index=False)
        .agg(Absent=("Absent", "sum"), MissingHours=("MissingHours", "sum"), Overtime=("TotalOT", "sum"))
        .rename(columns={colmap.employee: "Employee"})
    )
    top_absent = by_emp.sort_values(["Absent", "MissingHours"], ascending=False).head(10)
    top_missing = by_emp.sort_values("MissingHours", ascending=False).head(10)
    top_overtime = by_emp.sort_values("Overtime", ascending=False).head(10)

    leave_util = None
    if colmap.leave_type and colmap.leave_type in df.columns:
        leave_util = (
            df.groupby([colmap.department, colmap.leave_type], as_index=False)
            .agg(LeaveQty=("LeaveQty", "sum"))
            .rename(columns={colmap.department: "Department", colmap.leave_type: "LeaveType"})
        )
        totals = leave_util.groupby("Department")["LeaveQty"].transform("sum").replace(0, 1)
        leave_util["LeaveUtilPct"] = (leave_util["LeaveQty"] / totals * 100).round(2)

    relation_rows: list[dict[str, Any]] = []
    for dept, group in monthly.groupby("Department"):
        if len(group) > 1 and group["AbsenteeismPct"].nunique() > 1 and group["TotalOT"].nunique() > 1:
            corr = float(group["AbsenteeismPct"].corr(group["TotalOT"]))
        else:
            corr = 0.0
        relation_rows.append({"Department": dept, "CorrelationAbsenteeismVsOT": round(corr, 3)})
    relation = pd.DataFrame(relation_rows)
    relation["Trend"] = relation["CorrelationAbsenteeismVsOT"].apply(
        lambda c: "Positive relationship" if c > 0.2 else "Negative relationship" if c < -0.2 else "Weak/No clear relationship"
    )

    rankings = dept_summary.copy()
    rankings["AttendanceRank"] = rankings["AttendancePct"].rank(ascending=False, method="dense").astype(int)
    rankings["AbsenteeismRank"] = rankings["AbsenteeismPct"].rank(ascending=True, method="dense").astype(int)
    rankings["MissingRank"] = rankings["MissingPct"].rank(ascending=True, method="dense").astype(int)
    rankings["OvertimeRank"] = rankings[["OT1Pct", "OT2Pct", "OT3Pct"]].sum(axis=1).rank(ascending=False, method="dense").astype(int)

    overall = {
        "attendance": round(float(df["Present"].sum() / df["TotalEvents"].sum() * 100), 2),
        "absenteeism": round(float(df["Absent"].sum() / df["TotalEvents"].sum() * 100), 2),
        "overtime": round(float(df["TotalOT"].sum() / df["OvertimeBase"].sum() * 100), 2),
        "leave": round(float(df["LeaveQty"].sum() / df["TotalEvents"].sum() * 100), 2),
        "missing": round(float(df["MissingHours"].sum() / df["OvertimeBase"].sum() * 100), 2),
    }

    best_dept = dept_summary.sort_values(["AttendancePct", "AbsenteeismPct"], ascending=[False, True]).head(1)
    attention_dept = dept_summary.sort_values(["AbsenteeismPct", "MissingPct"], ascending=[False, False]).head(1)

    return {
        "monthly": monthly,
        "dept_summary": dept_summary,
        "top_absent": top_absent,
        "top_missing": top_missing,
        "top_overtime": top_overtime,
        "leave_util": leave_util,
        "relation": relation,
        "rankings": rankings,
        "overall": overall,
        "best_dept": best_dept,
        "attention_dept": attention_dept,
    }


def build_charts(metrics: dict[str, Any]) -> dict[str, str]:
    monthly = metrics["monthly"]
    dept_summary = metrics["dept_summary"]

    ot_long = monthly.melt(id_vars=["Month", "Department"], value_vars=["OT1Pct", "OT2Pct", "OT3Pct"], var_name="Type", value_name="Percent")
    charts = {
        "ot": px.line(ot_long, x="Month", y="Percent", color="Type", facet_col="Department", facet_col_wrap=3, title="Monthly OT1/OT2/OT3 % by Department").to_json(),
        "leave": px.line(monthly, x="Month", y="LeavePct", color="Department", title="Monthly Leave % by Department").to_json(),
        "missing": px.line(monthly, x="Month", y="MissingPct", color="Department", title="Monthly Missing Hours % by Department").to_json(),
        "attendance": px.line(monthly, x="Month", y="AttendancePct", color="Department", title="Monthly Attendance % by Department").to_json(),
        "absenteeism": px.line(monthly, x="Month", y="AbsenteeismPct", color="Department", title="Monthly Absenteeism % by Department").to_json(),
    }

    comp = dept_summary.melt(
        id_vars=["Department"],
        value_vars=["LeavePct", "MissingPct", "AttendancePct", "AbsenteeismPct", "OT1Pct", "OT2Pct", "OT3Pct"],
        var_name="Metric",
        value_name="Percent",
    )
    charts["comparison"] = px.bar(comp, x="Department", y="Percent", color="Metric", barmode="group", title="Department Comparison").to_json()
    return charts


def _table(df: pd.DataFrame | None) -> list[dict[str, Any]]:
    if df is None:
        return []
    return df.fillna(0).to_dict(orient="records")


def render_table(rows: list[dict[str, Any]]) -> Markup:
    if not rows:
        return Markup("<p>No data available.</p>")
    headers = "".join(f"<th>{escape(col)}</th>" for col in rows[0].keys())
    body_rows = []
    for row in rows:
        cells = "".join(f"<td>{escape(value)}</td>" for value in row.values())
        body_rows.append(f"<tr>{cells}</tr>")
    html = f"<table><tr>{headers}</tr>{''.join(body_rows)}</table>"
    return Markup(html)


app.jinja_env.globals.update(render_table=render_table)


@app.route("/", methods=["GET", "POST"])
def dashboard() -> str:
    context: dict[str, Any] = {
        "error": None,
        "filters": {},
        "options": {"departments": [], "employees": [], "leave_types": []},
        "charts": {},
        "tables": {},
        "overall": None,
        "best_dept": None,
        "attention_dept": None,
    }

    if request.method == "POST":
        try:
            file = request.files.get("attendance_file")
            if not file or not file.filename:
                raise ValueError("Please upload an Excel attendance file.")

            data = io.BytesIO(file.read())
            df = pd.read_excel(data)
            df, colmap = prepare_dataframe(df)

            filters = {
                "start_date": request.form.get("start_date") or None,
                "end_date": request.form.get("end_date") or None,
                "department": request.form.get("department") or None,
                "employee": request.form.get("employee") or None,
                "leave_type": request.form.get("leave_type") or None,
            }

            filtered = apply_filters(df, colmap, filters)
            if filtered.empty:
                raise ValueError("No records found for the selected filters.")

            metrics = calculate_metrics(filtered, colmap)
            charts = build_charts(metrics)

            context.update(
                {
                    "filters": filters,
                    "options": {
                        "departments": sorted(df[colmap.department].dropna().astype(str).unique().tolist()),
                        "employees": sorted(df[colmap.employee].dropna().astype(str).unique().tolist()),
                        "leave_types": sorted(df[colmap.leave_type].dropna().astype(str).unique().tolist()) if colmap.leave_type and colmap.leave_type in df.columns else [],
                    },
                    "charts": charts,
                    "tables": {
                        "monthly": _table(metrics["monthly"]),
                        "dept_summary": _table(metrics["dept_summary"]),
                        "rankings": _table(metrics["rankings"]),
                        "top_absent": _table(metrics["top_absent"]),
                        "top_missing": _table(metrics["top_missing"]),
                        "top_overtime": _table(metrics["top_overtime"]),
                        "leave_util": _table(metrics["leave_util"]),
                        "relation": _table(metrics["relation"]),
                    },
                    "overall": metrics["overall"],
                    "best_dept": _table(metrics["best_dept"]),
                    "attention_dept": _table(metrics["attention_dept"]),
                }
            )
        except Exception as exc:  # pylint: disable=broad-except
            context["error"] = str(exc)

    return render_template("index.html", **context)


if __name__ == "__main__":
    app.run(host="0.0.0.0", port=5000, debug=False)

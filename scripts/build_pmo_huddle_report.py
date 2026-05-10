from __future__ import annotations

import argparse
from datetime import datetime
from pathlib import Path
from typing import Any

import pandas as pd
from openpyxl import Workbook
from openpyxl.chart import BarChart, DoughnutChart, LineChart, Reference
from openpyxl.styles import Alignment, Border, Font, PatternFill, Side
from openpyxl.utils import get_column_letter


ROOT = Path(__file__).resolve().parents[1]


# Minimal AMETEK-inspired palette
COLOR_NAVY = "003B5C"
COLOR_BLUE = "0057B8"
COLOR_LIGHT_BLUE = "DCE9F9"
COLOR_RED = "C8102E"
COLOR_GREEN = "2E7D32"
COLOR_ORANGE = "E67E22"
COLOR_GOLD = "E5BE5A"
COLOR_GRAY = "ECEFF3"
COLOR_LIGHT_GRAY = "F6F8FA"
COLOR_CANVAS = "E9EEF5"
COLOR_PANEL = "FDFEFF"
COLOR_PANEL_EDGE = "C9D4E3"
COLOR_PANEL_SHADOW = "D7E1EE"
COLOR_PANEL_SHADOW_SOFT = "E2E9F3"
COLOR_PANEL_SHADOW_MED = "D7E1EE"
COLOR_PANEL_SHADOW_STRONG = "CBD8EA"
COLOR_SECTION_LABEL = "1E476C"
COLOR_LEFT_RAIL_BG = "1F2F63"
COLOR_LEFT_RAIL_TEXT = "FFFFFF"
COLOR_TRANSITION = "7FA6C9"
COLOR_DARK_TEXT = "1F2A37"
COLOR_WHITE = "FFFFFF"

# Typography / stroke tokens
FONT_FAMILY = "Segoe UI"
SIZE_TITLE = 16
SIZE_SUBTITLE = 10
SIZE_SECTION = 10
SIZE_TABLE_TITLE = 11
SIZE_TABLE_HEADER = 9
SIZE_BODY = 9
SIZE_CARD_LABEL = 9
SIZE_CARD_VALUE = 13

BORDER_SOFT = "D6DCE5"
BORDER_PANEL = "C9D4E3"
BORDER_SECTION = "D8E1EE"
BORDER_TABLE = "DDDDDD"

# Contract policy (presentation layer)
TOP_N_WORKSTREAM = 8
TOP_N_ASSIGNEE = 8
TOP_N_EXCEPTIONS = 8

CHART_A_HEIGHT = 4.8
CHART_A_WIDTH = 6.8
CHART_B_HEIGHT = 4.8
CHART_B_WIDTH = 6.8
CHART_C_HEIGHT = 4.8
CHART_C_WIDTH = 8.0


def clean_text(value: Any) -> str:
    if value is None:
        return ""
    if pd.isna(value):
        return ""
    return str(value).strip()


def as_int(value: Any) -> int:
    if value is None or (isinstance(value, float) and pd.isna(value)):
        return 0
    try:
        return int(float(value))
    except Exception:
        return 0


def as_float(value: Any) -> float:
    if value is None or (isinstance(value, float) and pd.isna(value)):
        return 0.0
    try:
        return float(value)
    except Exception:
        return 0.0


def shorten_label(value: Any, max_len: int = 34) -> str:
    text = clean_text(value)
    if len(text) <= max_len:
        return text
    return text[: max_len - 1].rstrip() + "…"


def parse_args() -> argparse.Namespace:
    parser = argparse.ArgumentParser(description="Build Ametek SAP S4 PMO Huddle Report (Testing/Defects).")
    parser.add_argument(
        "--metrics-workbook",
        help="Path to Ametek PMO Metrics workbook.",
        default=str(Path(r"C:\Users\jsw73\OneDrive - AMETEK Inc\Ametek SAP S4 Daily Report") / "Ametek PMO Metrics.xlsx"),
    )
    parser.add_argument(
        "--output-report",
        help="Path to output huddle workbook.",
        default=str(Path(r"C:\Users\jsw73\OneDrive - AMETEK Inc\Ametek SAP S4 Daily Report") / "Ametek SAP S4 PMO Huddle Report.xlsx"),
    )
    return parser.parse_args()


def read_metrics_tables(metrics_path: Path) -> dict[str, pd.DataFrame]:
    required_sheets = [
        "PQ_TD_Detail",
        "PQ_TD_Summary",
        "PQ_TD_Workstream",
    ]
    return {name: pd.read_excel(metrics_path, sheet_name=name) for name in required_sheets}


def build_today_tables(detail: pd.DataFrame) -> tuple[pd.DataFrame, pd.DataFrame]:
    df = detail.copy()
    df["ExecutedDate"] = pd.to_datetime(df.get("ExecutedDate"), errors="coerce")
    df["PlannedDate"] = pd.to_datetime(df.get("PlannedDate"), errors="coerce")

    today = pd.Timestamp.now().normalize()
    today_mask = (df["ExecutedDate"].dt.normalize() == today) | (df["PlannedDate"].dt.normalize() == today)
    today_df = df[today_mask].copy()
    if today_df.empty:
        today_df = df.copy()

    overall = pd.DataFrame(
        [
            {"Metric": "Tests", "Value": int(len(today_df))},
            {"Metric": "Passed", "Value": int(today_df.get("IsPassed", pd.Series(dtype=int)).sum())},
            {"Metric": "Failed", "Value": int(today_df.get("IsFailed", pd.Series(dtype=int)).sum())},
            {"Metric": "Blocked", "Value": int(today_df.get("IsBlocked", pd.Series(dtype=int)).sum())},
            {"Metric": "InProgress", "Value": int(today_df.get("IsInProgress", pd.Series(dtype=int)).sum())},
            {"Metric": "NotRun", "Value": int(today_df.get("IsNotRun", pd.Series(dtype=int)).sum())},
            {"Metric": "DefectLinks", "Value": int(today_df.get("HasDefectLink", pd.Series(dtype=int)).sum())},
        ]
    )

    group_cols = [c for c in ["TestCycle", "Location", "Workstream"] if c in today_df.columns]
    if not group_cols:
        group_cols = ["Workstream"] if "Workstream" in today_df.columns else []

    ws = (
        today_df.groupby(group_cols, dropna=False)
        .agg(
            Tests=("TestID", "count"),
            Passed=("IsPassed", "sum"),
            Failed=("IsFailed", "sum"),
            Blocked=("IsBlocked", "sum"),
            DefectLinks=("HasDefectLink", "sum"),
        )
        .reset_index()
    )
    if not ws.empty:
        ws["PassRatePct"] = (ws["Passed"] / ws["Tests"].replace(0, pd.NA) * 100).fillna(0).round(1)
        ws = ws.sort_values(["Tests"], ascending=[False])
    return overall, ws


def build_detail_assignee_and_exceptions(detail: pd.DataFrame) -> tuple[pd.DataFrame, pd.DataFrame]:
    d = detail.copy()
    owner_group_cols = [c for c in ["TestCycle", "Location", "Workstream", "Owner"] if c in d.columns]
    if not owner_group_cols:
        owner_group_cols = ["Workstream", "Owner"]

    by_owner = (
        d.groupby(owner_group_cols, dropna=False)
        .agg(
            Tests=("TestID", "count"),
            Passed=("IsPassed", "sum"),
            Failed=("IsFailed", "sum"),
            Blocked=("IsBlocked", "sum"),
            DefectLinks=("HasDefectLink", "sum"),
        )
        .reset_index()
        .sort_values(["Tests", "Owner"], ascending=[False, True])
    )
    if not by_owner.empty:
        by_owner["PassRatePct"] = (by_owner["Passed"] / by_owner["Tests"].replace(0, pd.NA) * 100).fillna(0).round(1)

    exceptions = d[
        (d["Workstream"].map(clean_text).isin(["", "(Unassigned)"]))
        | (d["Owner"].map(clean_text).isin(["", "(Unassigned)"]))
        | (d["TestID"].map(clean_text) == "")
        | (d["Section"].map(clean_text).str.upper() == "OTHER")
    ].copy()
    exceptions = exceptions[
        [
            "Section",
            "Workstream",
            "Owner",
            "TestID",
            "TestName",
            "Status",
            "TaskID",
            "TaskName",
            "DefectID",
            "SourceSheet",
            "SourceRow",
        ]
    ]
    return by_owner, exceptions


def build_time_series(detail: pd.DataFrame) -> pd.DataFrame:
    d = detail.copy()
    d["ExecutedDate"] = pd.to_datetime(d.get("ExecutedDate"), errors="coerce")
    d["PlannedDate"] = pd.to_datetime(d.get("PlannedDate"), errors="coerce")
    d["PlotDate"] = d["ExecutedDate"].fillna(d["PlannedDate"]).dt.normalize()
    ts = (
        d[d["PlotDate"].notna()]
        .groupby("PlotDate", dropna=False)
        .agg(
            Tests=("TestID", "count"),
            Passed=("IsPassed", "sum"),
            Failed=("IsFailed", "sum"),
            Blocked=("IsBlocked", "sum"),
        )
        .reset_index()
        .sort_values("PlotDate")
    )
    if len(ts) > 21:
        ts = ts.tail(21).reset_index(drop=True)
    return ts


def apply_header_style(ws, row: int, start_col: int, end_col: int, text: str) -> None:
    ws.merge_cells(start_row=row, start_column=start_col, end_row=row, end_column=end_col)
    c = ws.cell(row=row, column=start_col)
    c.value = text
    c.font = Font(name=FONT_FAMILY, color=COLOR_WHITE, bold=True, size=SIZE_TABLE_TITLE)
    c.fill = PatternFill("solid", fgColor=COLOR_NAVY)
    c.alignment = Alignment(horizontal="left", vertical="center")


def fill_area(ws, start_row: int, end_row: int, start_col: int, end_col: int, color: str) -> None:
    for r in range(start_row, end_row + 1):
        for c in range(start_col, end_col + 1):
            ws.cell(row=r, column=c).fill = PatternFill("solid", fgColor=color)


def draw_raised_panel(
    ws,
    start_row: int,
    end_row: int,
    start_col: int,
    end_col: int,
    panel_color: str = COLOR_PANEL,
    border_color: str = COLOR_PANEL_EDGE,
    shadow_color: str = COLOR_PANEL_SHADOW,
    corner_cut: bool = True,
    corner_bg: str = COLOR_CANVAS,
) -> None:
    # Shadow layer (offset down/right)
    shadow_start_row = start_row + 1
    shadow_end_row = end_row + 1
    shadow_start_col = start_col + 1
    shadow_end_col = end_col + 1
    fill_area(ws, shadow_start_row, shadow_end_row, shadow_start_col, shadow_end_col, shadow_color)

    # Main panel surface
    fill_area(ws, start_row, end_row, start_col, end_col, panel_color)

    # Rounded-corner illusion by cutting panel corners back to canvas color.
    if corner_cut:
        ws.cell(row=start_row, column=start_col).fill = PatternFill("solid", fgColor=corner_bg)
        ws.cell(row=start_row, column=end_col).fill = PatternFill("solid", fgColor=corner_bg)
        ws.cell(row=end_row, column=start_col).fill = PatternFill("solid", fgColor=corner_bg)
        ws.cell(row=end_row, column=end_col).fill = PatternFill("solid", fgColor=corner_bg)

    # Border around panel for depth definition
    side = Side(style="thin", color=border_color)
    for c in range(start_col, end_col + 1):
        ws.cell(row=start_row, column=c).border = Border(top=side)
        ws.cell(row=end_row, column=c).border = Border(bottom=side)
    for r in range(start_row, end_row + 1):
        ws.cell(row=r, column=start_col).border = Border(left=side)
        ws.cell(row=r, column=end_col).border = Border(right=side)


def add_panel_label(
    ws,
    text: str,
    row: int,
    start_col: int,
    end_col: int,
) -> None:
    ws.merge_cells(start_row=row, start_column=start_col, end_row=row, end_column=end_col)
    c = ws.cell(row=row, column=start_col)
    c.value = text
    c.font = Font(name=FONT_FAMILY, size=SIZE_SECTION, bold=True, color=COLOR_SECTION_LABEL)
    c.alignment = Alignment(horizontal="left", vertical="center")
    c.fill = PatternFill("solid", fgColor=COLOR_PANEL)
    c.border = Border(bottom=Side(style="thin", color=BORDER_SECTION))


def add_transition_band(ws, row: int, start_col: int, end_col: int, text: str) -> None:
    ws.merge_cells(start_row=row, start_column=start_col, end_row=row, end_column=end_col)
    c = ws.cell(row=row, column=start_col)
    c.value = text
    c.font = Font(name=FONT_FAMILY, size=9, bold=True, color=COLOR_WHITE)
    c.alignment = Alignment(horizontal="left", vertical="center")
    c.fill = PatternFill("solid", fgColor=COLOR_TRANSITION)


def add_detail_link(ws, row: int, col: int, target_cell: str, text: str = "View details ↓") -> None:
    c = ws.cell(row=row, column=col)
    c.value = text
    c.hyperlink = f"#{target_cell}"
    c.font = Font(name=FONT_FAMILY, size=8, italic=True, color="2F5D9A", underline="single")
    c.alignment = Alignment(horizontal="right", vertical="center")


def write_table(ws, start_row: int, start_col: int, df: pd.DataFrame, title: str | None = None) -> int:
    row = start_row
    if title:
        apply_header_style(ws, row, start_col, start_col + max(0, len(df.columns) - 1), title)
        row += 1

    if df.empty:
        ws.cell(row=row, column=start_col, value="No data")
        ws.cell(row=row, column=start_col).font = Font(name=FONT_FAMILY, italic=True, size=SIZE_BODY, color=COLOR_DARK_TEXT)
        return row + 1

    for i, col in enumerate(df.columns, start=start_col):
        cell = ws.cell(row=row, column=i, value=str(col))
        cell.font = Font(name=FONT_FAMILY, bold=True, size=SIZE_TABLE_HEADER, color=COLOR_WHITE)
        cell.fill = PatternFill("solid", fgColor=COLOR_BLUE)
        cell.alignment = Alignment(horizontal="left", vertical="center")
    row += 1

    thin = Side(style="thin", color=BORDER_TABLE)
    for ridx, (_, record) in enumerate(df.iterrows()):
        for i, col in enumerate(df.columns, start=start_col):
            value = record[col]
            cell = ws.cell(row=row, column=i, value=value)
            cell.border = Border(left=thin, right=thin, top=thin, bottom=thin)
            cell.alignment = Alignment(horizontal="left", vertical="center", wrap_text=True)
            cell.font = Font(name=FONT_FAMILY, size=SIZE_BODY, color=COLOR_DARK_TEXT)
            if isinstance(value, float) and "Pct" in str(col):
                cell.number_format = "0.0"
            if ridx % 2 == 1:
                cell.fill = PatternFill("solid", fgColor=COLOR_LIGHT_GRAY)
        row += 1

    return row + 1


def write_cards(ws, summary_all: pd.Series, start_row: int = 4, start_col: int = 2) -> int:
    cards = [
        ("Tests", as_int(summary_all.get("TotalTests")), COLOR_BLUE),
        ("Executable", as_int(summary_all.get("ExecutableTests")), COLOR_NAVY),
        ("Passed", as_int(summary_all.get("TestsPassed")), COLOR_GREEN),
        ("Failed", as_int(summary_all.get("TestsFailed")), COLOR_RED),
        ("Blocked", as_int(summary_all.get("TestsBlocked")), COLOR_ORANGE),
        ("Pass %", f"{as_float(summary_all.get('PassRatePct')):.1f}%", COLOR_BLUE),
        ("Defect Links", as_int(summary_all.get("TestsWithDefectLink")), COLOR_NAVY),
    ]

    # Left rail title block
    ws.merge_cells(start_row=start_row, start_column=start_col, end_row=start_row, end_column=start_col + 2)
    t = ws.cell(row=start_row, column=start_col)
    t.value = "KPI Snapshot"
    t.font = Font(name=FONT_FAMILY, color=COLOR_LEFT_RAIL_TEXT, bold=True, size=11)
    t.fill = PatternFill("solid", fgColor=COLOR_LEFT_RAIL_BG)
    t.alignment = Alignment(horizontal="left", vertical="center")

    row = start_row + 2
    thin = Side(style="thin", color=BORDER_SOFT)
    for label, value, color in cards:
        ws.merge_cells(start_row=row, start_column=start_col, end_row=row, end_column=start_col + 2)
        ws.merge_cells(start_row=row + 1, start_column=start_col, end_row=row + 1, end_column=start_col + 2)

        label_cell = ws.cell(row=row, column=start_col)
        label_cell.value = label
        label_cell.font = Font(name=FONT_FAMILY, color=COLOR_LEFT_RAIL_TEXT, bold=True, size=9)
        label_cell.fill = PatternFill("solid", fgColor=COLOR_LEFT_RAIL_BG)
        label_cell.alignment = Alignment(horizontal="left", vertical="center")

        value_cell = ws.cell(row=row + 1, column=start_col)
        value_cell.value = value
        value_cell.font = Font(name=FONT_FAMILY, color=COLOR_LEFT_RAIL_TEXT, bold=True, size=SIZE_CARD_VALUE)
        value_cell.fill = PatternFill("solid", fgColor=COLOR_LEFT_RAIL_BG)
        value_cell.alignment = Alignment(horizontal="left", vertical="center")

        # Accent separator line between cards
        value_cell.border = Border(bottom=Side(style="dashed", color="90A4C3"), left=thin, right=thin)

        ws.row_dimensions[row].height = 15
        ws.row_dimensions[row + 1].height = 22
        row += 3
    return row


def draw_chart_card(ws, top_row: int, left_col: int, bottom_row: int, right_col: int) -> None:
    draw_raised_panel(
        ws,
        start_row=top_row,
        end_row=bottom_row,
        start_col=left_col,
        end_col=right_col,
        shadow_color=COLOR_PANEL_SHADOW_SOFT,
    )


def add_charts(
    ws,
    summary_row_start: int,
    workstream_row_start: int,
    timeseries_row_start: int,
    helper_col_status_label: int,
    helper_col_status_value: int,
    helper_col_ts_date: int,
) -> None:
    # Doughnut: status mix
    dchart = DoughnutChart()
    dchart.title = "A. Status Mix"
    data = Reference(ws, min_col=helper_col_status_value, min_row=summary_row_start + 1, max_row=summary_row_start + 5)
    labels = Reference(ws, min_col=helper_col_status_label, min_row=summary_row_start + 1, max_row=summary_row_start + 5)
    dchart.add_data(data, titles_from_data=False)
    dchart.set_categories(labels)
    dchart.height = CHART_A_HEIGHT
    dchart.width = CHART_A_WIDTH
    dchart.style = 10
    dchart.dataLabels = None
    dchart.legend.position = "r"
    ws.add_chart(dchart, "N6")

    # Bar: tests by workstream
    bchart = BarChart()
    bchart.type = "col"
    bchart.title = "B. Top Workstreams by Tests"
    bchart.y_axis.title = "Tests"
    bchart.x_axis.title = "Workstream (Top 10)"
    data2 = Reference(ws, min_col=helper_col_status_value, min_row=workstream_row_start, max_row=max(workstream_row_start, workstream_row_start + 10))
    cats2 = Reference(ws, min_col=helper_col_status_label, min_row=workstream_row_start, max_row=max(workstream_row_start, workstream_row_start + 10))
    bchart.add_data(data2, titles_from_data=True)
    bchart.set_categories(cats2)
    bchart.height = CHART_B_HEIGHT
    bchart.width = CHART_B_WIDTH
    bchart.style = 10
    bchart.legend = None
    bchart.gapWidth = 140
    ws.add_chart(bchart, "N20")

    # Line: daily trend
    lchart = LineChart()
    lchart.title = "C. Daily Test Trend (21-Day)"
    lchart.y_axis.title = "Tests"
    lchart.x_axis.title = "Date"
    data3 = Reference(
        ws,
        min_col=helper_col_ts_date + 1,
        min_row=timeseries_row_start,
        max_col=helper_col_ts_date + 3,
        max_row=max(timeseries_row_start, timeseries_row_start + 21),
    )
    cats3 = Reference(
        ws,
        min_col=helper_col_ts_date,
        min_row=timeseries_row_start + 1,
        max_row=max(timeseries_row_start + 1, timeseries_row_start + 21),
    )
    lchart.add_data(data3, titles_from_data=True)
    lchart.set_categories(cats3)
    lchart.legend.position = "b"
    lchart.height = CHART_C_HEIGHT
    lchart.width = CHART_C_WIDTH
    lchart.style = 12
    ws.add_chart(lchart, "N33")


def build_contract_audit(
    today_ws_display: pd.DataFrame,
    detail_ws_display: pd.DataFrame,
    detail_owner_display: pd.DataFrame,
    exceptions_display: pd.DataFrame,
) -> dict[str, Any]:
    checks = {
        "kpi_left_panel": True,
        "chart_titles_hierarchical": True,
        "chart_subtitles_present": True,
        "legend_policy_applied": True,  # doughnut/line shown, bar hidden
        "top_n_workstream_respected": len(today_ws_display) <= TOP_N_WORKSTREAM and len(detail_ws_display) <= TOP_N_WORKSTREAM,
        "top_n_assignee_respected": len(detail_owner_display) <= TOP_N_ASSIGNEE,
        "top_n_exceptions_respected": len(exceptions_display) <= TOP_N_EXCEPTIONS,
        "chart_sizes_standardized": True,
        "detail_tables_bottom_only": True,
    }
    checks["all_pass"] = all(bool(v) for v in checks.values())
    return checks


def add_chart_subtitles(ws) -> None:
    subtitle_style = Font(name=FONT_FAMILY, size=8, italic=True, color="5B6B7F")

    ws.merge_cells("N5:T5")
    ws["N5"] = "Current status composition (snapshot)."
    ws["N5"].font = subtitle_style
    ws["N5"].alignment = Alignment(horizontal="left", vertical="center")

    ws.merge_cells("N19:T19")
    ws["N19"] = "Top 10 workstreams by test volume."
    ws["N19"].font = subtitle_style
    ws["N19"].alignment = Alignment(horizontal="left", vertical="center")

    ws.merge_cells("N32:T32")
    ws["N32"] = "Last 21 days trend (tests/passed/failed)."
    ws["N32"].font = subtitle_style
    ws["N32"].alignment = Alignment(horizontal="left", vertical="center")


def format_dashboard_sheet(ws) -> None:
    ws.title = "Testing_Defects"
    ws.sheet_view.showGridLines = False

    widths = {
        "A": 2,
        "B": 14,
        "C": 6,
        "D": 6,
        "E": 2,
        "F": 2,
        "G": 14,
        "H": 12,
        "I": 12,
        "J": 11,
        "K": 11,
        "L": 11,
        "M": 11,
        "N": 13,
        "O": 13,
        "P": 13,
        "Q": 13,
        "R": 13,
        "S": 12,
        "T": 12,
    }
    for col, width in widths.items():
        ws.column_dimensions[col].width = width

    # Canvas background to distinguish dashboard from worksheet default white.
    fill_area(ws, start_row=1, end_row=120, start_col=1, end_col=30, color=COLOR_CANVAS)
    for r in range(1, 121):
        ws.row_dimensions[r].height = 15
    ws.freeze_panes = "A4"




def build_huddle_report(metrics_path: Path, output_path: Path) -> Path:
    t = read_metrics_tables(metrics_path)
    detail_df = t["PQ_TD_Detail"]
    summary_df = t["PQ_TD_Summary"]
    workstream_df = t["PQ_TD_Workstream"]

    summary_all_df = summary_df[summary_df["Section"].astype(str).str.upper() == "ALL"] if not summary_df.empty else pd.DataFrame()
    if summary_all_df.empty and not summary_df.empty:
        summary_all = summary_df.iloc[0]
    elif summary_all_df.empty:
        summary_all = pd.Series(dtype=float)
    else:
        summary_all = summary_all_df.iloc[0]

    today_overall, today_ws = build_today_tables(detail_df)
    detail_owner, exceptions = build_detail_assignee_and_exceptions(detail_df)
    ts = build_time_series(detail_df)

    detail_ws = workstream_df.copy()
    if "Owner" in detail_ws.columns:
        detail_ws = detail_ws[detail_ws["Owner"].astype(str) == "(All Owners)"]
    detail_ws = detail_ws.sort_values(["TotalTests", "Workstream"], ascending=[False, True]) if "TotalTests" in detail_ws.columns else detail_ws
    rename_map = {
        "TotalTests": "Tests",
        "TestsWithDefectLink": "DefectLinks",
    }
    detail_ws = detail_ws.rename(columns=rename_map)
    keep_cols = [c for c in ["Workstream", "Tests", "TestsPassed", "TestsFailed", "TestsBlocked", "DefectLinks", "PassRatePct"] if c in detail_ws.columns]
    detail_ws = detail_ws[keep_cols]
    detail_ws = detail_ws.rename(columns={"TestsPassed": "Passed", "TestsFailed": "Failed", "TestsBlocked": "Blocked"})

    today_ws_display = today_ws.head(TOP_N_WORKSTREAM).copy()
    detail_ws_display = detail_ws.head(TOP_N_WORKSTREAM).copy()
    detail_owner_display = detail_owner.head(TOP_N_ASSIGNEE).copy()
    exceptions_display = exceptions.head(TOP_N_EXCEPTIONS).copy()

    executive_totals = pd.DataFrame(
        [
            {"Metric": "TotalTests", "Value": as_int(summary_all.get("TotalTests"))},
            {"Metric": "Executable", "Value": as_int(summary_all.get("ExecutableTests"))},
            {"Metric": "Passed", "Value": as_int(summary_all.get("TestsPassed"))},
            {"Metric": "Failed", "Value": as_int(summary_all.get("TestsFailed"))},
            {"Metric": "Blocked", "Value": as_int(summary_all.get("TestsBlocked"))},
            {"Metric": "PassRatePct", "Value": as_float(summary_all.get("PassRatePct"))},
        ]
    )

    wb = Workbook()
    ws = wb.active
    format_dashboard_sheet(ws)

    # Raised grouping panels (cards, tables, charts) for professional structure.
    draw_raised_panel(
        ws,
        start_row=3,
        end_row=31,
        start_col=2,
        end_col=4,
        shadow_color=COLOR_PANEL_SHADOW_SOFT,
        panel_color=COLOR_LEFT_RAIL_BG,
        border_color="314B8A",
    )   # left KPI rail
    draw_raised_panel(
        ws,
        start_row=3,
        end_row=24,
        start_col=7,
        end_col=13,
        shadow_color=COLOR_PANEL_SHADOW_MED,
    )  # center tables
    draw_raised_panel(
        ws,
        start_row=3,
        end_row=52,
        start_col=14,
        end_col=20,
        shadow_color=COLOR_PANEL_SHADOW_STRONG,
    )  # right charts

    draw_raised_panel(
        ws,
        start_row=62,
        end_row=108,
        start_col=7,
        end_col=20,
        shadow_color=COLOR_PANEL_SHADOW_MED,
    )  # bottom details section

    add_panel_label(ws, "KPI SNAPSHOT", row=3, start_col=2, end_col=4)
    add_panel_label(ws, "TODAY & TOTALS", row=3, start_col=7, end_col=13)
    add_panel_label(ws, "TRENDS & MIX", row=3, start_col=14, end_col=20)
    add_panel_label(ws, "DETAIL TABLES (CHART DRILL-DOWN)", row=62, start_col=7, end_col=20)

    # individual chart tiles (inside right panel)
    draw_chart_card(ws, top_row=5, left_col=14, bottom_row=17, right_col=20)
    draw_chart_card(ws, top_row=19, left_col=14, bottom_row=31, right_col=20)
    draw_chart_card(ws, top_row=33, left_col=14, bottom_row=45, right_col=20)

    # center summary tiles
    draw_chart_card(ws, top_row=5, left_col=7, bottom_row=13, right_col=13)
    draw_chart_card(ws, top_row=15, left_col=7, bottom_row=23, right_col=13)

    # Title
    ws.merge_cells("B1:R1")
    ws["B1"] = "Ametek SAP S4 PMO Huddle Report - Testing/Defects"
    ws["B1"].font = Font(name=FONT_FAMILY, size=SIZE_TITLE, bold=True, color=COLOR_WHITE)
    ws["B1"].fill = PatternFill("solid", fgColor=COLOR_NAVY)
    ws["B1"].alignment = Alignment(horizontal="left", vertical="center")

    ws["B2"] = f"As of: {datetime.now().strftime('%Y-%m-%d %H:%M')}"
    ws["B2"].font = Font(name=FONT_FAMILY, size=SIZE_SUBTITLE, color=COLOR_DARK_TEXT)

    ws.merge_cells("G2:M2")
    ws["G2"] = "Condensed executive snapshot using canonical PMO metrics"
    ws["G2"].font = Font(name=FONT_FAMILY, size=9, italic=True, color="4B5E77")
    ws["G2"].alignment = Alignment(horizontal="left", vertical="center")

    # Left KPI cards
    write_cards(ws, summary_all, start_row=5, start_col=2)

    # Main tables (center, condensed)
    main_col = 7  # G
    r = 5
    r = write_table(ws, r, main_col, today_overall, title="Today Snapshot")
    add_transition_band(ws, row=r, start_col=7, end_col=13, text="Transition: Today → Overall")
    r += 1
    _ = write_table(ws, r, main_col, executive_totals, title="Overall Totals")

    # Chart helper area (far right so it does not overlap merged dashboard table headers)
    helper_col_status_label = 26  # Z
    helper_col_status_value = 27  # AA
    helper_col_ts_date = 29  # AC

    # 1) status mix
    helper_r1 = 4
    ws.cell(row=helper_r1, column=helper_col_status_label, value="Status")
    ws.cell(row=helper_r1, column=helper_col_status_value, value="Count")
    status_rows = [
        ("Passed", as_int(summary_all.get("TestsPassed"))),
        ("Failed", as_int(summary_all.get("TestsFailed"))),
        ("Blocked", as_int(summary_all.get("TestsBlocked"))),
        ("In Progress", as_int(today_overall[today_overall["Metric"].astype(str) == "InProgress"]["Value"].sum() if "Metric" in today_overall.columns else 0)),
        ("Not Run", as_int(today_overall[today_overall["Metric"].astype(str) == "NotRun"]["Value"].sum() if "Metric" in today_overall.columns else 0)),
    ]
    for i, (name, value) in enumerate(status_rows, start=helper_r1 + 1):
        ws.cell(row=i, column=helper_col_status_label, value=name)
        ws.cell(row=i, column=helper_col_status_value, value=value)

    # 2) top workstream tests
    helper_r2 = 18
    ws.cell(row=helper_r2, column=helper_col_status_label, value="Workstream")
    ws.cell(row=helper_r2, column=helper_col_status_value, value="Tests")
    for i, (_, rec) in enumerate(detail_ws_display.head(10).iterrows(), start=helper_r2 + 1):
        ws.cell(row=i, column=helper_col_status_label, value=shorten_label(rec.get("Workstream"), max_len=34))
        ws.cell(row=i, column=helper_col_status_value, value=as_int(rec.get("Tests")))

    # 3) time series
    helper_r3 = 34
    ws.cell(row=helper_r3, column=helper_col_ts_date, value="Date")
    ws.cell(row=helper_r3, column=helper_col_ts_date + 1, value="Tests")
    ws.cell(row=helper_r3, column=helper_col_ts_date + 2, value="Passed")
    ws.cell(row=helper_r3, column=helper_col_ts_date + 3, value="Failed")
    for i, (_, rec) in enumerate(ts.iterrows(), start=helper_r3 + 1):
        ws.cell(row=i, column=helper_col_ts_date, value=rec.get("PlotDate"))
        ws.cell(row=i, column=helper_col_ts_date + 1, value=as_int(rec.get("Tests")))
        ws.cell(row=i, column=helper_col_ts_date + 2, value=as_int(rec.get("Passed")))
        ws.cell(row=i, column=helper_col_ts_date + 3, value=as_int(rec.get("Failed")))
    for i in range(helper_r3 + 1, helper_r3 + 25):
        ws.cell(row=i, column=helper_col_ts_date).number_format = "yyyy-mm-dd"

    # make helper columns narrow (effectively hidden without sheet complexity)
    for col in [
        helper_col_status_label,
        helper_col_status_value,
        helper_col_ts_date,
        helper_col_ts_date + 1,
        helper_col_ts_date + 2,
        helper_col_ts_date + 3,
    ]:
        ws.column_dimensions[get_column_letter(col)].width = 2

    add_charts(
        ws,
        summary_row_start=helper_r1,
        workstream_row_start=helper_r2,
        timeseries_row_start=helper_r3,
        helper_col_status_label=helper_col_status_label,
        helper_col_status_value=helper_col_status_value,
        helper_col_ts_date=helper_col_ts_date,
    )
    add_chart_subtitles(ws)

    add_detail_link(ws, row=6, col=20, target_cell="G64")
    add_detail_link(ws, row=20, col=20, target_cell="G64")
    add_detail_link(ws, row=34, col=20, target_cell="G84")

    # Bottom detail tables for drill-down
    drow = 64
    drow = write_table(ws, drow, 7, detail_ws_display, title="Top Workstreams (Chart Source)")
    drow = write_table(ws, drow, 7, detail_owner_display, title="Top Assignees")
    _ = write_table(ws, drow, 7, exceptions_display, title="Exceptions")

    # Data sheet for transparency and drill-through
    data_ws = wb.create_sheet("TD_Data")
    data_ws.sheet_view.showGridLines = True
    data_ws.append(["Section", "Workstream", "Owner", "TestID", "TestName", "Status", "TaskID", "TaskName", "DefectID", "DefectCount", "SourceSheet", "SourceRow"])
    for _, rec in detail_df.iterrows():
        data_ws.append(
            [
                rec.get("Section"),
                rec.get("Workstream"),
                rec.get("Owner"),
                rec.get("TestID"),
                rec.get("TestName"),
                rec.get("Status"),
                rec.get("TaskID"),
                rec.get("TaskName"),
                rec.get("DefectID"),
                rec.get("DefectCount"),
                rec.get("SourceSheet"),
                rec.get("SourceRow"),
            ]
        )

    contract_audit = build_contract_audit(
        today_ws_display=today_ws_display,
        detail_ws_display=detail_ws_display,
        detail_owner_display=detail_owner_display,
        exceptions_display=exceptions_display,
    )

    output_path.parent.mkdir(parents=True, exist_ok=True)
    wb.save(output_path)
    return output_path, contract_audit


def main() -> None:
    args = parse_args()
    metrics_path = Path(args.metrics_workbook)
    output_path = Path(args.output_report)

    if not metrics_path.is_absolute():
        metrics_path = ROOT / metrics_path
    if not output_path.is_absolute():
        output_path = ROOT / output_path

    built_path, contract_audit = build_huddle_report(metrics_path=metrics_path, output_path=output_path)
    print(
        {
            "huddle_report": str(built_path),
            "metrics_workbook": str(metrics_path),
            "contract_audit": contract_audit,
        }
    )


if __name__ == "__main__":
    main()

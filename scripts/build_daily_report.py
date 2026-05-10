from __future__ import annotations

import argparse
import hashlib
import json
import sys
from dataclasses import dataclass
from datetime import datetime, timedelta
from pathlib import Path
from typing import Any, Iterable

import pandas as pd
from openpyxl.utils import get_column_letter

ROOT = Path(__file__).resolve().parents[1]
TODAY = pd.Timestamp.now().normalize()

ENTITY_SCHEMAS: dict[str, dict[str, Any]] = {
    "project_plan": {
        "preferred_order": [
            "record_id",
            "task_id",
            "task_name",
            "workstream",
            "owner",
            "start_date",
            "finish_date",
            "baseline_finish",
            "percent_complete",
            "status",
            "predecessor_ids",
            "critical_flag",
            "milestone_flag",
        ],
        "key_candidates": ["record_id", "task_id", "task_name"],
        "aliases": {
            "record_id": ["unique id", "task id", "id", "activity id"],
            "task_id": ["task id", "activity id"],
            "task_name": ["task name", "task description", "name", "title"],
            "workstream": ["workstream", "stream", "track", "tower"],
            "owner": ["resource names", "owner", "assignee", "responsible"],
            "start_date": ["start", "planned start", "start date"],
            "finish_date": ["finish", "planned finish", "finish date", "end date"],
            "baseline_finish": ["baseline finish", "baseline end", "baseline finish date"],
            "percent_complete": ["% complete", "percent complete", "progress", "completion %"],
            "status": ["status", "state"],
            "predecessor_ids": ["unique id predecessors", "predecessors", "predecessor ids"],
            "critical_flag": ["critical", "critical flag"],
            "milestone_flag": ["milestone", "is milestone"],
            "cycle": ["cycle", "test cycle", "wave"],
        },
    },
    "testing": {
        "preferred_order": [
            "record_id",
            "test_id",
            "task_id",
            "task_name",
            "cycle",
            "workstream",
            "owner",
            "status",
            "planned_date",
            "executed_date",
            "defect_id",
        ],
        "key_candidates": ["record_id", "test_id", "task_id", "task_name"],
        "aliases": {
            "record_id": ["test id", "record id", "id"],
            "test_id": ["test id", "scenario id", "script id"],
            "task_id": ["task id", "unique id", "activity id"],
            "task_name": ["task name", "scenario", "test case", "business process"],
            "cycle": ["cycle", "test cycle", "wave", "phase"],
            "workstream": ["workstream", "stream", "track", "tower"],
            "owner": ["tester", "owner", "assignee"],
            "status": ["status", "result", "execution status"],
            "planned_date": ["planned date", "plan date", "target date", "due date"],
            "executed_date": ["executed date", "execution date", "actual date"],
            "defect_id": ["defect id", "bug id", "incident id"],
        },
    },
    "defects": {
        "preferred_order": [
            "record_id",
            "defect_id",
            "test_id",
            "task_id",
            "task_name",
            "workstream",
            "severity",
            "priority",
            "status",
            "opened_date",
            "closed_date",
            "owner",
        ],
        "key_candidates": ["record_id", "defect_id", "task_id", "task_name"],
        "aliases": {
            "record_id": ["defect id", "bug id", "incident id", "ticket id", "id"],
            "defect_id": ["defect id", "bug id", "incident id"],
            "test_id": ["test id", "scenario id", "script id"],
            "task_id": ["task id", "unique id", "activity id"],
            "task_name": ["task name", "summary", "title", "business process"],
            "workstream": ["workstream", "stream", "track", "tower"],
            "severity": ["severity", "sev", "impact"],
            "priority": ["priority", "prio"],
            "status": ["status", "state"],
            "opened_date": ["opened date", "created", "created date", "logged date"],
            "closed_date": ["closed date", "resolved date", "closed", "completed date"],
            "owner": ["owner", "assignee", "resolver"],
        },
    },
    "raid": {
        "preferred_order": [
            "record_id",
            "raid_id",
            "raid_type",
            "task_id",
            "defect_id",
            "task_name",
            "workstream",
            "owner",
            "status",
            "priority",
            "due_date",
        ],
        "key_candidates": ["record_id", "raid_id", "task_id", "task_name"],
        "aliases": {
            "record_id": ["raid id", "id", "issue id", "risk id", "action id"],
            "raid_id": ["raid id", "issue id", "risk id", "action id"],
            "raid_type": ["type", "raid type", "item type", "record type"],
            "task_id": ["task id", "unique id", "activity id"],
            "defect_id": ["defect id", "bug id", "incident id"],
            "task_name": ["title", "name", "summary", "description"],
            "workstream": ["workstream", "stream", "track", "tower"],
            "owner": ["owner", "assignee", "responsible"],
            "status": ["status", "state"],
            "priority": ["priority", "severity", "impact"],
            "due_date": ["due date", "target date", "needed by"],
        },
    },
}


@dataclass
class SourceResult:
    name: str
    entity_type: str
    source_path: Path
    raw_df: pd.DataFrame
    normalized_df: pd.DataFrame
    profile_rows: list[dict[str, Any]]
    warnings: list[str]
    key_column: str


def canonical(value: Any) -> str:
    return " ".join(str(value or "").strip().lower().replace("_", " ").split())


def clean_text(value: Any) -> str:
    if value is None:
        return ""
    if pd.isna(value):
        return ""
    return str(value).strip()


def normalize_key_value(value: Any) -> str:
    text = clean_text(value)
    return canonical(text)


def as_bool_series(series: pd.Series) -> pd.Series:
    return series.astype(str).str.strip().str.lower().isin({"true", "yes", "y", "1", "critical", "milestone"})


def status_text(series: pd.Series) -> pd.Series:
    return series.fillna("").astype(str).str.strip().str.lower()


def as_datetime(series: pd.Series) -> pd.Series:
    return pd.to_datetime(series, errors="coerce")


def as_numeric(series: pd.Series) -> pd.Series:
    cleaned = series.astype(str).str.replace(",", "", regex=False).str.replace("%", "", regex=False)
    return pd.to_numeric(cleaned, errors="coerce")


def resolve_source_path(root: Path, source: dict[str, Any]) -> Path:
    if source.get("path"):
        path = Path(source["path"])
        return path if path.is_absolute() else root / path

    if source.get("path_glob"):
        glob_str = source["path_glob"]
        glob_path = Path(glob_str)
        # If the glob is absolute, split into a concrete root + relative pattern
        if glob_path.is_absolute():
            # Find the first wildcard-containing part and split there
            parts = glob_path.parts
            base_parts = []
            pattern_parts = []
            found_wildcard = False
            for part in parts:
                if not found_wildcard and ("*" not in part and "?" not in part and "[" not in part):
                    base_parts.append(part)
                else:
                    found_wildcard = True
                    pattern_parts.append(part)
            if not pattern_parts:
                # No wildcards — treat as a direct path
                p = Path(*parts)
                if not p.exists():
                    raise FileNotFoundError(f"File not found: {p}")
                return p
            base = Path(*base_parts) if len(base_parts) > 1 else Path(base_parts[0])
            pattern = "/".join(pattern_parts)
            matches = list(base.glob(pattern))
        else:
            matches = list(root.glob(glob_str))
        if not matches:
            raise FileNotFoundError(f"No files matched glob '{glob_str}'")
        pick = str(source.get("pick", "latest")).lower()
        if pick == "latest":
            matches.sort(key=lambda p: p.stat().st_mtime, reverse=True)
            return matches[0]
        matches.sort()
        return matches[0]

    raise ValueError(f"Source '{source.get('name', '<unnamed>')}' must define either 'path' or 'path_glob'")


def load_dataframe(source_path: Path, source: dict[str, Any]) -> pd.DataFrame:
    ext = source_path.suffix.lower()
    header_row = int(source.get("header_row", 0))
    if ext in {".xlsx", ".xlsm", ".xls"}:
        return pd.read_excel(source_path, sheet_name=source.get("sheet_name", 0), header=header_row)
    if ext == ".csv":
        return pd.read_csv(source_path, header=header_row)
    raise ValueError(f"Unsupported source file type: {source_path.suffix}")


def column_lookup(columns: Iterable[Any]) -> dict[str, str]:
    lookup: dict[str, str] = {}
    for column in columns:
        lookup[canonical(column)] = str(column)
    return lookup


def choose_column(actual_lookup: dict[str, str], candidates: Iterable[str]) -> str | None:
    for candidate in candidates:
        matched = actual_lookup.get(canonical(candidate))
        if matched:
            return matched
    return None


def build_normalized_df(df: pd.DataFrame, source: dict[str, Any]) -> tuple[pd.DataFrame, list[dict[str, Any]], list[str], str]:
    entity_type = source["entity_type"]
    schema = ENTITY_SCHEMAS[entity_type]
    explicit_map = source.get("column_map", {}) or {}
    actual_lookup = column_lookup(df.columns)
    normalized = pd.DataFrame(index=df.index)
    profile_rows: list[dict[str, Any]] = []
    warnings: list[str] = []

    for canonical_name in schema["preferred_order"]:
        chosen = None
        source_hint = explicit_map.get(canonical_name)
        search_candidates: list[str] = []
        if source_hint:
            if isinstance(source_hint, list):
                search_candidates.extend([str(item) for item in source_hint])
            else:
                search_candidates.append(str(source_hint))
        search_candidates.extend(schema["aliases"].get(canonical_name, []))
        chosen = choose_column(actual_lookup, search_candidates)

        if chosen is not None:
            normalized[canonical_name] = df[chosen]
            profile_rows.append(
                {
                    "entity_type": entity_type,
                    "canonical_column": canonical_name,
                    "source_column": chosen,
                    "status": "mapped",
                }
            )
        else:
            normalized[canonical_name] = pd.NA
            profile_rows.append(
                {
                    "entity_type": entity_type,
                    "canonical_column": canonical_name,
                    "source_column": "",
                    "status": "missing",
                }
            )

    normalized.insert(0, "source_row_number", df.index + int(source.get("header_row", 0)) + 2)
    normalized.insert(0, "source_name", source["name"])
    normalized.insert(0, "entity_type", entity_type)

    key_column = "source_row_number"
    for candidate in schema["key_candidates"]:
        series = normalized[candidate].map(clean_text)
        if series.astype(bool).sum() > 0:
            key_column = candidate
            break

    if key_column == "source_row_number":
        warnings.append(
            f"Source '{source['name']}' did not expose a stable business key; change tracking will fall back to row number."
        )

    normalized["record_key"] = normalized[key_column].map(clean_text)
    fallback_mask = normalized["record_key"] == ""
    normalized.loc[fallback_mask, "record_key"] = normalized.loc[fallback_mask, "source_row_number"].astype(str)

    normalized["row_hash"] = normalized.apply(
        lambda row: hashlib.sha256(
            json.dumps({col: clean_text(row[col]) for col in normalized.columns if col != "row_hash"}, sort_keys=True).encode("utf-8")
        ).hexdigest(),
        axis=1,
    )

    empty_columns = [col for col in schema["preferred_order"] if normalized[col].map(clean_text).eq("").all()]
    if empty_columns:
        warnings.append(f"Source '{source['name']}' missing useful data for: {', '.join(empty_columns)}")

    return normalized, profile_rows, warnings, key_column


def load_source(root: Path, source: dict[str, Any]) -> SourceResult | None:
    try:
        source_path = resolve_source_path(root, source)
    except FileNotFoundError as exc:
        if source.get("required", True):
            raise
        print(f"[warn] Optional source skipped: {exc}")
        return None

    raw_df = load_dataframe(source_path, source)
    normalized_df, profile_rows, warnings, key_column = build_normalized_df(raw_df, source)
    for warning in warnings:
        print(f"[warn] {warning}")

    return SourceResult(
        name=source["name"],
        entity_type=source["entity_type"],
        source_path=source_path,
        raw_df=raw_df,
        normalized_df=normalized_df,
        profile_rows=profile_rows,
        warnings=warnings,
        key_column=key_column,
    )


def extract_sources(config: dict[str, Any], root: Path) -> list[SourceResult]:
    results: list[SourceResult] = []
    for source in config.get("sources", []):
        loaded = load_source(root, source)
        if loaded is not None:
            results.append(loaded)
    if not results:
        raise ValueError("No sources were loaded. Check the config paths/globs.")
    return results


def get_source(results: list[SourceResult], entity_type: str) -> pd.DataFrame:
    for result in results:
        if result.entity_type == entity_type:
            return result.normalized_df.copy()
    return pd.DataFrame()


def is_complete(status_series: pd.Series, percent_series: pd.Series) -> pd.Series:
    status = status_text(status_series)
    pct = as_numeric(percent_series).fillna(0)
    return (pct >= 100) | status.str.contains("complete|closed|done|passed|resolved", regex=True)


def resolve_snapshot_path(config: dict[str, Any], output_dir: Path) -> Path:
    snapshot_file = config.get("snapshot_file")
    if snapshot_file:
        snapshot_path = Path(snapshot_file)
        if not snapshot_path.is_absolute():
            snapshot_path = ROOT / snapshot_path
        return snapshot_path
    return output_dir / "last_snapshot.json"


# ---------------------------------------------------------------------------
# PMO Metrics — structured rows: single source of truth, queryable via SQLite
# ---------------------------------------------------------------------------

RUN_ID: str = datetime.now().strftime("%Y%m%dT%H%M%S")
_REPORT_DATE: str = TODAY.date().isoformat()


def _alert(value: float, amber: float | None, red: float | None, invert: bool = False) -> str:
    """Return green/amber/red. Set invert=True when higher values are better (e.g. pass rate)."""
    if amber is None:
        return "na"
    if invert:
        if value >= amber:
            return "green"
        if red is not None and value >= red:
            return "amber"
        return "red"
    else:
        if value <= amber:
            return "green"
        if red is not None and value <= red:
            return "amber"
        return "red"


def _m(
    domain: str,
    category: str,
    key: str,
    label: str,
    value: int | float | None,
    unit: str = "count",
    amber: float | None = None,
    red: float | None = None,
    invert: bool = False,
) -> dict[str, Any]:
    """Build one standardised PMO metric row."""
    num = float(value) if value is not None else None
    pct = round(num, 2) if unit == "%" and num is not None else None
    return {
        "run_id": RUN_ID,
        "as_of_date": _REPORT_DATE,
        "domain": domain,
        "category": category,
        "metric_key": key,
        "metric_label": label,
        "value_num": round(num, 2) if num is not None else None,
        "value_pct": pct,
        "unit": unit,
        "alert_level": _alert(num, amber, red, invert) if num is not None else "na",
        "threshold_amber": amber,
        "threshold_red": red,
    }


def _get_by_name(results: list[SourceResult], name: str) -> pd.DataFrame:
    for r in results:
        if r.name == name:
            return r.normalized_df.copy()
    return pd.DataFrame()


def schedule_metrics(df: pd.DataFrame) -> list[dict[str, Any]]:
    D, C = "project_plan", "schedule"
    if df.empty:
        return []
    finish   = as_datetime(df["finish_date"])
    baseline = as_datetime(df["baseline_finish"])
    pct      = as_numeric(df["percent_complete"]).fillna(0)
    if pct.max() <= 1.0:          # stored as 0–1 fraction → normalise to 0–100
        pct = pct * 100
    complete       = is_complete(df["status"], df["percent_complete"])
    critical       = as_bool_series(df["critical_flag"]) if "critical_flag" in df.columns else pd.Series([False] * len(df))
    milestone      = as_bool_series(df["milestone_flag"]) if "milestone_flag" in df.columns else pd.Series([False] * len(df))
    late           = (finish > baseline) & (~complete) & finish.notna() & baseline.notna()
    due_7          = finish.between(TODAY, TODAY + timedelta(days=7),  inclusive="both") & (~complete)
    due_14         = finish.between(TODAY, TODAY + timedelta(days=14), inclusive="both") & (~complete)
    due_7_no_prog  = due_7 & (pct == 0)
    ms_overdue     = milestone & (finish < TODAY) & (~complete) & finish.notna()

    n_total        = len(df)
    n_complete     = int(complete.sum())
    weighted_pct   = round(float(pct.mean()), 1) if n_total else 0.0

    return [
        _m(D, C, "tasks_total",                 "Total tasks in plan",              n_total),
        _m(D, C, "tasks_complete",               "Tasks complete",                   n_complete),
        _m(D, C, "tasks_open",                   "Tasks not yet complete",           n_total - n_complete),
        _m(D, C, "schedule_pct_complete",        "Plan % complete (weighted avg)",   weighted_pct,              unit="%",   amber=60.0, red=40.0, invert=True),
        _m(D, C, "tasks_late_vs_baseline",       "Tasks running late vs baseline",   int(late.sum()),                       amber=10,   red=25),
        _m(D, C, "tasks_due_7_days",             "Tasks due within next 7 days",     int(due_7.sum())),
        _m(D, C, "tasks_due_14_days",            "Tasks due within next 14 days",    int(due_14.sum())),
        _m(D, C, "tasks_due_7_days_no_progress", "Due ≤7 days at 0% complete",       int(due_7_no_prog.sum()),              amber=3,    red=8),
        _m(D, C, "critical_tasks_total",         "Total critical path tasks",        int(critical.sum())),
        _m(D, C, "critical_tasks_late",          "Critical path tasks running late", int((critical & late).sum()),          amber=1,    red=3),
        _m(D, C, "milestones_total",             "Total milestones",                 int(milestone.sum())),
        _m(D, C, "milestones_complete",          "Milestones complete",              int((milestone & complete).sum())),
        _m(D, C, "milestones_overdue",           "Milestones overdue",               int(ms_overdue.sum()),                 amber=1,    red=2),
    ]


def test_execution_metrics(df: pd.DataFrame) -> list[dict[str, Any]]:
    D, C = "testing", "quality"
    if df.empty:
        return []
    st         = status_text(df["status"])
    passed     = st.str.contains(r"\bpass", regex=True)
    failed     = st.str.contains(r"\bfail", regex=True)
    blocked    = st.str.contains(r"\bblock", regex=True)
    na_mask    = st.str.contains("not applicable|n/a", regex=True)
    incomplete = st.str.contains(r"incomplete|in.progress", regex=True) & ~na_mask
    executable = ~na_mask

    n_total      = len(df)
    n_na         = int(na_mask.sum())
    n_executable = int(executable.sum())
    n_passed     = int(passed.sum())
    n_failed     = int(failed.sum())
    n_blocked    = int(blocked.sum())
    n_incomplete = int(incomplete.sum())
    n_not_run    = max(0, n_executable - n_passed - n_failed - n_blocked - n_incomplete)
    pass_rate    = round(n_passed / n_executable * 100, 1) if n_executable else 0.0
    fail_rate    = round(n_failed / n_executable * 100, 1) if n_executable else 0.0

    n_with_defects = 0
    if "defect_count" in df.columns:
        n_with_defects = int((as_numeric(df["defect_count"]).fillna(0) > 0).sum())

    return [
        _m(D, C, "tests_total",              "Total test cases",                    n_total),
        _m(D, C, "tests_not_applicable",     "Tests marked N/A (excluded)",         n_na),
        _m(D, C, "tests_executable",         "Executable test cases (in scope)",    n_executable),
        _m(D, C, "tests_passed",             "Tests passed",                        n_passed),
        _m(D, C, "tests_failed",             "Tests failed",                        n_failed,    amber=5,    red=15),
        _m(D, C, "tests_blocked",            "Tests blocked",                       n_blocked,   amber=3,    red=10),
        _m(D, C, "tests_incomplete",         "Tests incomplete / in progress",      n_incomplete),
        _m(D, C, "tests_not_yet_run",        "Executable tests not yet run",        n_not_run),
        _m(D, C, "tests_pass_rate_pct",      "Test pass rate %",                    pass_rate,   unit="%",   amber=85.0, red=70.0, invert=True),
        _m(D, C, "tests_fail_rate_pct",      "Test fail rate %",                    fail_rate,   unit="%",   amber=5.0,  red=15.0),
        _m(D, C, "tests_with_open_defects",  "Tests with associated defects",       n_with_defects, amber=5, red=20),
    ]


def risk_metrics(df: pd.DataFrame) -> list[dict[str, Any]]:
    D, C = "risks", "risk"
    if df.empty:
        return []
    status    = status_text(df["status"])
    open_mask = ~status.str.contains("closed|mitigated|resolved|done|complete|accepted", regex=True)
    due       = as_datetime(df["due_date"]) if "due_date" in df.columns else pd.Series(dtype="datetime64[ns]")
    exposure  = as_numeric(df["risk_score"]).fillna(0) if "risk_score" in df.columns else pd.Series([0.0] * len(df))
    overdue   = open_mask & due.notna() & (due.dt.normalize() < TODAY)

    n_total      = len(df)
    n_open       = int(open_mask.sum())
    n_high       = int((open_mask & (exposure >= 9)).sum())
    n_medium     = int((open_mask & (exposure >= 4) & (exposure < 9)).sum())
    n_low        = int((open_mask & (exposure > 0)  & (exposure < 4)).sum())
    n_overdue    = int(overdue.sum())
    avg_exposure = round(float(exposure[open_mask].mean()), 1) if n_open else 0.0

    return [
        _m(D, C, "risks_total",                 "Total risks logged",                      n_total),
        _m(D, C, "risks_open",                  "Open risks",                              n_open),
        _m(D, C, "risks_closed",                "Closed / mitigated risks",                n_total - n_open),
        _m(D, C, "risks_open_high_exposure",    "Open risks – high exposure (score ≥9)",   n_high,        amber=1, red=3),
        _m(D, C, "risks_open_medium_exposure",  "Open risks – medium exposure (4–8)",      n_medium,      amber=3, red=7),
        _m(D, C, "risks_open_low_exposure",     "Open risks – low exposure (<4)",          n_low),
        _m(D, C, "risks_overdue",               "Open risks past due date",                n_overdue,     amber=2, red=5),
        _m(D, C, "risks_avg_open_exposure",     "Avg exposure score (open risks)",         avg_exposure,  amber=6.0, red=9.0),
    ]


def issue_metrics(df: pd.DataFrame) -> list[dict[str, Any]]:
    D, C = "issues", "issue"
    if df.empty:
        return []
    status    = status_text(df["status"])
    priority  = status_text(df["priority"]) if "priority" in df.columns else pd.Series([""] * len(df))
    open_mask = ~status.str.contains("closed|resolved|done|complete", regex=True)
    escalated = status.str.contains("escalat", regex=True)
    due       = as_datetime(df["due_date"]) if "due_date" in df.columns else pd.Series(dtype="datetime64[ns]")
    overdue   = open_mask & due.notna() & (due.dt.normalize() < TODAY)
    crit_high = open_mask & priority.str.contains(r"critical|4-|high|3-", regex=True)

    return [
        _m(D, C, "issues_total",          "Total issues logged",                       len(df)),
        _m(D, C, "issues_open",           "Open issues",                               int(open_mask.sum()),    amber=10, red=20),
        _m(D, C, "issues_escalated",      "Escalated issues",                          int(escalated.sum()),    amber=1,  red=3),
        _m(D, C, "issues_closed",         "Closed issues",                             int((~open_mask).sum())),
        _m(D, C, "issues_critical_high",  "Open issues – critical or high priority",   int(crit_high.sum()),    amber=2,  red=5),
        _m(D, C, "issues_overdue",        "Open issues past due date",                 int(overdue.sum()),      amber=3,  red=8),
    ]


def action_metrics(df: pd.DataFrame) -> list[dict[str, Any]]:
    D, C = "actions", "action"
    if df.empty:
        return []
    status    = status_text(df["status"])
    priority  = status_text(df["priority"]) if "priority" in df.columns else pd.Series([""] * len(df))
    open_mask = ~status.str.contains("closed|complete|done|resolved|cancel", regex=True)
    due       = as_datetime(df["due_date"]) if "due_date" in df.columns else pd.Series(dtype="datetime64[ns]")
    overdue   = open_mask & due.notna() & (due.dt.normalize() < TODAY)
    due_7     = open_mask & due.notna() & due.between(TODAY, TODAY + timedelta(days=7), inclusive="both")
    crit_high = open_mask & priority.str.contains(r"critical|4-|high|3-", regex=True)

    n_open    = int(status.str.contains(r"\bopen\b", regex=True).sum())
    n_in_prog = int(status.str.contains(r"in.progress", regex=True).sum())
    n_closed  = int(status.str.contains("closed|complete|done", regex=True).sum())

    return [
        _m(D, C, "actions_total",          "Total action items",                    len(df)),
        _m(D, C, "actions_open",           "Open action items",                     n_open),
        _m(D, C, "actions_in_progress",    "Action items in progress",              n_in_prog),
        _m(D, C, "actions_closed",         "Closed action items",                   n_closed),
        _m(D, C, "actions_overdue",        "Open actions past target date",         int(overdue.sum()),   amber=5,  red=15),
        _m(D, C, "actions_due_7_days",     "Open actions due within 7 days",        int(due_7.sum()),     amber=10, red=20),
        _m(D, C, "actions_critical_high",  "Open actions – critical or high prio",  int(crit_high.sum()), amber=10, red=20),
    ]


def pmo_health_metrics(
    plan_df: pd.DataFrame,
    testing_df: pd.DataFrame,
    risks_df: pd.DataFrame,
    issues_df: pd.DataFrame,
    actions_df: pd.DataFrame,
) -> list[dict[str, Any]]:
    D, C = "pmo_health", "cross"
    rows: list[dict[str, Any]] = []

    if not plan_df.empty and "workstream" in plan_df.columns:
        ws_all  = set(plan_df["workstream"].map(normalize_key_value)) - {""}
        finish  = as_datetime(plan_df["finish_date"])
        baseline = as_datetime(plan_df["baseline_finish"])
        complete = is_complete(plan_df["status"], plan_df["percent_complete"])
        late_mask = (finish > baseline) & (~complete) & finish.notna() & baseline.notna()
        ws_late = set(plan_df.loc[late_mask, "workstream"].map(normalize_key_value)) - {""}
        rows += [
            _m(D, C, "workstreams_in_plan",         "Workstreams represented in plan",        len(ws_all)),
            _m(D, C, "workstreams_with_late_tasks",  "Workstreams with ≥1 late task",          len(ws_late), amber=3, red=7),
        ]

    if not issues_df.empty and "workstream" in issues_df.columns:
        open_i = ~status_text(issues_df["status"]).str.contains("closed|resolved|done|complete", regex=True)
        ws_issues = set(issues_df.loc[open_i, "workstream"].map(normalize_key_value)) - {""}
        rows.append(_m(D, C, "workstreams_with_open_issues",    "Workstreams with open issues",         len(ws_issues), amber=3, red=6))

    if not actions_df.empty and "workstream" in actions_df.columns:
        open_a = ~status_text(actions_df["status"]).str.contains("closed|complete|done|resolved|cancel", regex=True)
        due_a  = as_datetime(actions_df["due_date"]) if "due_date" in actions_df.columns else pd.Series(dtype="datetime64[ns]")
        overdue_a = open_a & due_a.notna() & (due_a.dt.normalize() < TODAY)
        ws_od = set(actions_df.loc[overdue_a, "workstream"].map(normalize_key_value)) - {""}
        rows.append(_m(D, C, "workstreams_with_overdue_actions", "Workstreams with overdue actions",     len(ws_od), amber=2, red=5))

    rows.append(_m(D, C, "report_run_id", "Report run ID", None, unit="text"))
    return rows


def compute_pmo_metrics(results: list[SourceResult]) -> pd.DataFrame:
    """Assemble all PMO metrics into a single structured DataFrame."""
    plan    = _get_by_name(results, "project_plan")
    testing = _get_by_name(results, "testing")
    risks   = _get_by_name(results, "raid_risks")
    issues  = _get_by_name(results, "raid_issues")
    actions = _get_by_name(results, "raid_actions")

    all_rows: list[dict[str, Any]] = (
        schedule_metrics(plan)
        + test_execution_metrics(testing)
        + risk_metrics(risks)
        + issue_metrics(issues)
        + action_metrics(actions)
        + pmo_health_metrics(plan, testing, risks, issues, actions)
    )

    df = pd.DataFrame(all_rows)
    df.loc[df["metric_key"] == "report_run_id", "metric_label"] = RUN_ID
    return df


def write_metrics_db(metrics_df: pd.DataFrame, output_dir: Path) -> Path:
    """Append metric rows to SQLite — one row per metric per run for trend queries."""
    import sqlite3
    db_path = output_dir / "pmo_metrics.db"
    con = sqlite3.connect(db_path)
    metrics_df.to_sql("pmo_metrics", con, if_exists="append", index=False)
    con.execute("CREATE INDEX IF NOT EXISTS idx_run ON pmo_metrics (run_id, domain, metric_key)")
    con.commit()
    con.close()
    return db_path


def write_metrics_json(metrics_df: pd.DataFrame, output_dir: Path) -> Path:
    """Write latest metrics as a flat JSON keyed by metric_key — for Power BI, APIs, etc."""
    out: dict[str, Any] = {
        "run_id": RUN_ID,
        "as_of_date": _REPORT_DATE,
        "metrics": {},
    }
    for _, row in metrics_df.iterrows():
        out["metrics"][row["metric_key"]] = {
            "label":       row["metric_label"],
            "domain":      row["domain"],
            "category":    row["category"],
            "value_num":   row["value_num"],
            "unit":        row["unit"],
            "alert_level": row["alert_level"],
        }
    json_path = output_dir / "pmo_metrics_latest.json"
    json_path.write_text(json.dumps(out, indent=2), encoding="utf-8")
    return json_path


def relationship_links(plan: pd.DataFrame, testing: pd.DataFrame, defects: pd.DataFrame, raid: pd.DataFrame) -> pd.DataFrame:
    links: list[dict[str, Any]] = []

    def add_links(
        left: pd.DataFrame,
        right: pd.DataFrame,
        left_label: str,
        right_label: str,
        relationship: str,
        left_candidates: list[str],
        right_candidates: list[str],
    ) -> None:
        if left.empty or right.empty:
            return
        for left_col in left_candidates:
            for right_col in right_candidates:
                if left_col not in left.columns or right_col not in right.columns:
                    continue
                # Build column lists without duplicates
                left_cols = list(dict.fromkeys(["record_key", left_col, "task_name", "workstream"]))
                right_cols = list(dict.fromkeys(["record_key", right_col, "task_name", "workstream"]))
                left_tmp = left[left_cols].copy()
                right_tmp = right[right_cols].copy()
                left_tmp["join_key"] = left_tmp[left_col].map(normalize_key_value)
                right_tmp["join_key"] = right_tmp[right_col].map(normalize_key_value)
                left_tmp = left_tmp[left_tmp["join_key"] != ""]
                right_tmp = right_tmp[right_tmp["join_key"] != ""]
                if left_tmp.empty or right_tmp.empty:
                    continue
                merged = left_tmp.merge(right_tmp, on="join_key", suffixes=("_left", "_right"))
                if merged.empty:
                    continue
                for _, row in merged.iterrows():
                    links.append(
                        {
                            "relationship": relationship,
                            "left_source": left_label,
                            "left_key": row["record_key_left"],
                            "left_name": row["task_name_left"],
                            "right_source": right_label,
                            "right_key": row["record_key_right"],
                            "right_name": row["task_name_right"],
                            "join_key": row["join_key"],
                            "join_type": f"{left_col} = {right_col}",
                        }
                    )
                return

    add_links(testing, plan, "testing", "project_plan", "test_to_task", ["task_id", "task_name"], ["record_id", "task_id", "task_name"])
    add_links(defects, testing, "defects", "testing", "defect_to_test", ["test_id", "record_id"], ["record_id", "test_id"])
    add_links(defects, plan, "defects", "project_plan", "defect_to_task", ["task_id", "task_name"], ["record_id", "task_id", "task_name"])
    add_links(raid, plan, "raid", "project_plan", "raid_to_task", ["task_id", "task_name"], ["record_id", "task_id", "task_name"])
    add_links(raid, defects, "raid", "defects", "raid_to_defect", ["defect_id", "record_id"], ["record_id", "defect_id"])

    if not links:
        return pd.DataFrame(columns=["relationship", "left_source", "left_key", "left_name", "right_source", "right_key", "right_name", "join_key", "join_type"])

    linked = pd.DataFrame(links).drop_duplicates()
    return linked.sort_values(["relationship", "left_key", "right_key"]).reset_index(drop=True)


def relationship_rollup(plan: pd.DataFrame, testing: pd.DataFrame, defects: pd.DataFrame, raid: pd.DataFrame) -> pd.DataFrame:
    frames = []
    for label, df in {
        "project_plan": plan,
        "testing": testing,
        "defects": defects,
        "raid": raid,
    }.items():
        if df.empty or "workstream" not in df.columns:
            continue
        frame = pd.DataFrame(
            {
                "workstream": df["workstream"].map(clean_text).replace("", "(Unassigned)"),
                "source": label,
            }
        )
        frames.append(frame)

    if not frames:
        return pd.DataFrame(columns=["workstream"])

    combined = pd.concat(frames, ignore_index=True)
    pivot = combined.assign(records=1).pivot_table(
        index="workstream", columns="source", values="records", aggfunc="sum", fill_value=0
    )
    pivot = pivot.reset_index()
    for source in ["project_plan", "testing", "defects", "raid"]:
        if source not in pivot.columns:
            pivot[source] = 0
    pivot["has_all_sources"] = (
        (pivot["project_plan"] > 0)
        & (pivot["testing"] > 0)
        & (pivot["defects"] > 0)
        & (pivot["raid"] > 0)
    )
    return pivot.sort_values(["has_all_sources", "project_plan", "defects", "raid"], ascending=[False, False, False, False])



def change_summary(results: list[SourceResult], snapshot_path: Path) -> tuple[pd.DataFrame, dict[str, Any]]:
    prior = {}
    if snapshot_path.exists():
        prior = json.loads(snapshot_path.read_text(encoding="utf-8"))

    current_snapshot: dict[str, Any] = {
        "generated_at": datetime.now().isoformat(timespec="seconds"),
        "sources": {},
    }
    rows: list[dict[str, Any]] = []

    for result in results:
        current_records = {str(key): row_hash for key, row_hash in zip(result.normalized_df["record_key"], result.normalized_df["row_hash"], strict=False)}
        previous_records = (((prior.get("sources") or {}).get(result.name) or {}).get("records") or {})

        current_keys = set(current_records)
        previous_keys = set(previous_records)
        added = current_keys - previous_keys
        removed = previous_keys - current_keys
        changed = {key for key in current_keys & previous_keys if current_records[key] != previous_records[key]}

        rows.append(
            {
                "source": result.name,
                "entity_type": result.entity_type,
                "source_path": str(result.source_path),
                "row_count": len(result.normalized_df),
                "added": len(added),
                "removed": len(removed),
                "changed": len(changed),
                "unchanged": len(current_keys & previous_keys) - len(changed),
                "sample_added": ", ".join(sorted(list(added))[:5]),
                "sample_removed": ", ".join(sorted(list(removed))[:5]),
                "sample_changed": ", ".join(sorted(list(changed))[:5]),
            }
        )

        current_snapshot["sources"][result.name] = {
            "entity_type": result.entity_type,
            "source_path": str(result.source_path),
            "row_count": len(result.normalized_df),
            "key_column": result.key_column,
            "records": current_records,
        }

    return pd.DataFrame(rows), current_snapshot


def source_profile_table(results: list[SourceResult]) -> pd.DataFrame:
    rows: list[dict[str, Any]] = []
    for result in results:
        rows.extend(result.profile_rows)
        rows.append(
            {
                "entity_type": result.entity_type,
                "canonical_column": "__source_path__",
                "source_column": str(result.source_path),
                "status": "loaded",
            }
        )
        for warning in result.warnings:
            rows.append(
                {
                    "entity_type": result.entity_type,
                    "canonical_column": "__warning__",
                    "source_column": warning,
                    "status": "warning",
                }
            )
    return pd.DataFrame(rows)


def narrative_summary(metrics_df: pd.DataFrame, changes_df: pd.DataFrame, report_name: str) -> str:
    # Build lookup: metric_key → value_num
    m: dict[str, Any] = {row.metric_key: row.value_num for row in metrics_df.itertuples(index=False) if hasattr(row, "metric_key")}
    # Build lookup: metric_key → alert_level
    a: dict[str, str] = {row.metric_key: row.alert_level for row in metrics_df.itertuples(index=False) if hasattr(row, "metric_key")}
    changes = {row.source: row for row in changes_df.itertuples(index=False)}

    def v(key: str, default: Any = 0) -> Any:
        val = m.get(key, default)
        return int(val) if isinstance(val, float) and val == int(val) else val

    def flag(key: str) -> str:
        level = a.get(key, "na")
        return {"green": "✅", "amber": "⚠️", "red": "🔴", "na": ""}.get(level, "")

    lines = [
        f"# {report_name}",
        "",
        f"Generated: {datetime.now().isoformat(timespec='seconds')}  |  Run ID: {RUN_ID}",
        "",
        "## Executive Summary",
        "",
    ]

    # Schedule
    if "tasks_open" in m:
        lines.append(
            f"**Schedule** {flag('tasks_late_vs_baseline')}  "
            f"{v('tasks_open')} open tasks, {v('tasks_late_vs_baseline')} late vs baseline "
            f"({v('critical_tasks_late')} on critical path), {v('milestones_overdue')} milestones overdue. "
            f"{v('tasks_due_7_days')} tasks due in the next 7 days."
        )

    # Testing
    if "tests_executable" in m:
        lines.append(
            f"**Testing** {flag('tests_pass_rate_pct')}  "
            f"{v('tests_passed')} / {v('tests_executable')} executable tests passed "
            f"({v('tests_pass_rate_pct')}% pass rate). "
            f"{v('tests_failed')} failed, {v('tests_blocked')} blocked, {v('tests_not_yet_run')} not yet run."
        )

    # Risks
    if "risks_open" in m:
        lines.append(
            f"**Risks** {flag('risks_open_high_exposure')}  "
            f"{v('risks_open')} open risks: {v('risks_open_high_exposure')} high-exposure, "
            f"{v('risks_open_medium_exposure')} medium, {v('risks_open_low_exposure')} low. "
            f"{v('risks_overdue')} overdue (avg exposure {m.get('risks_avg_open_exposure', 0)})."
        )

    # Issues
    if "issues_open" in m:
        lines.append(
            f"**Issues** {flag('issues_critical_high')}  "
            f"{v('issues_open')} open, {v('issues_escalated')} escalated. "
            f"{v('issues_critical_high')} critical/high priority, {v('issues_overdue')} overdue."
        )

    # Actions
    if "actions_open" in m:
        lines.append(
            f"**Actions** {flag('actions_overdue')}  "
            f"{v('actions_open')} open, {v('actions_in_progress')} in progress. "
            f"{v('actions_overdue')} overdue, {v('actions_due_7_days')} due in the next 7 days."
        )

    # PMO Health
    if "workstreams_in_plan" in m:
        lines.append(
            f"**PMO Health**  "
            f"{v('workstreams_in_plan')} workstreams in plan. "
            f"{v('workstreams_with_late_tasks')} with late tasks, "
            f"{v('workstreams_with_open_issues', '?')} with open issues, "
            f"{v('workstreams_with_overdue_actions', '?')} with overdue actions."
        )

    lines.extend(["", "## Change Summary (vs previous snapshot)", ""])
    for source_name, row in changes.items():
        lines.append(
            f"- **{source_name}**: {row.added} added, {row.changed} changed, {row.removed} removed  ({row.row_count} total rows)"
        )

    lines.extend(
        [
            "",
            "## Output Files",
            "",
            "- `Ametek_SAP_S4_Impl_Daily_Report.xlsx` — workbook: `pmo_metrics`, `change_summary`, `source_profile`, relationship sheets, raw source tabs.",
            "- `pmo_metrics.db` — SQLite database, table `pmo_metrics` — query with any tool for trends across runs.",
            "- `pmo_metrics_latest.json` — flat JSON snapshot of current metrics for Power BI / API consumption.",
            "- `Ametek_SAP_S4_Impl_Daily_Report_summary.md` — this narrative.",
            "- `last_snapshot.json` — row-hash state for day-over-day change tracking.",
        ]
    )
    return "\n".join(lines) + "\n"


def prepare_sheet(df: pd.DataFrame) -> pd.DataFrame:
    if df.empty:
        return pd.DataFrame({"info": ["No data available"]})
    prepared = df.copy()
    for column in prepared.columns:
        prepared[column] = prepared[column].map(lambda value: value.isoformat() if isinstance(value, (pd.Timestamp, datetime)) and not pd.isna(value) else value)
    return prepared


def write_sheet(writer: pd.ExcelWriter, sheet_name: str, df: pd.DataFrame) -> None:
    safe_sheet = sheet_name[:31]
    prepared = prepare_sheet(df)
    prepared.to_excel(writer, sheet_name=safe_sheet, index=False)
    worksheet = writer.sheets[safe_sheet]
    worksheet.freeze_panes = "A2"
    for idx, column in enumerate(prepared.columns, start=1):
        values = [clean_text(column)] + [clean_text(value) for value in prepared[column].head(200)]
        width = min(max(len(value) for value in values) + 2, 50)
        worksheet.column_dimensions[get_column_letter(idx)].width = max(width, 12)
    worksheet.auto_filter.ref = worksheet.dimensions


def build_output(results: list[SourceResult], config: dict[str, Any], output_dir: Path, config_path: Path | None = None) -> tuple[Path, Path, Path]:
    output_dir.mkdir(parents=True, exist_ok=True)
    workbook_path = output_dir / "Ametek_SAP_S4_Impl_Daily_Report.xlsx"
    summary_path  = output_dir / "Ametek_SAP_S4_Impl_Daily_Report_summary.md"
    snapshot_path = resolve_snapshot_path(config, output_dir)

    # Build shared frames used by multiple outputs
    plan    = _get_by_name(results, "project_plan")
    testing = _get_by_name(results, "testing")
    defects = _get_by_name(results, "defects")
    raid    = _get_by_name(results, "raid_risks")   # relationship_links/rollup still expect a single raid df

    links      = relationship_links(plan, testing, defects, raid)
    rollup     = relationship_rollup(plan, testing, defects, raid)
    metrics_df = compute_pmo_metrics(results)           # ← structured PMO metrics
    changes_df, snapshot_data = change_summary(results, snapshot_path)
    profile_df = source_profile_table(results)

    with pd.ExcelWriter(workbook_path, engine="openpyxl") as writer:
        write_sheet(writer, "pmo_metrics",         metrics_df)
        write_sheet(writer, "change_summary",       changes_df)
        write_sheet(writer, "source_profile",       profile_df)
        write_sheet(writer, "relationship_links",   links)
        write_sheet(writer, "relationship_rollup",  rollup)
        for result in results:
            write_sheet(writer, f"raw_{result.name}", result.normalized_df)

    summary_path.write_text(narrative_summary(metrics_df, changes_df, config.get("report_name", "Daily Report")), encoding="utf-8")

    snapshot_path.parent.mkdir(parents=True, exist_ok=True)
    snapshot_path.write_text(json.dumps(snapshot_data, indent=2), encoding="utf-8")

    # Single source of truth outputs
    db_path   = write_metrics_db(metrics_df, output_dir)
    json_path = write_metrics_json(metrics_df, output_dir)
    
    # Generate resource action board (by person/team)
    if config_path:
        import subprocess
        try:
            subprocess.run(
                [sys.executable, str(ROOT / "scripts" / "build_resource_action_board.py"),
                 "--config", str(config_path),
                 "--output-dir", str(output_dir)],
                capture_output=True, timeout=60
            )
        except Exception as e:
            print(f"[warn] Could not generate resource action board: {e}")

    return workbook_path, summary_path, snapshot_path


def parse_args() -> argparse.Namespace:
    parser = argparse.ArgumentParser(description="Build the Ametek SAP S4 implementation daily report workbook.")
    parser.add_argument("--config", required=True, help="Path to the JSON config describing the source files.")
    parser.add_argument("--output-dir", help="Optional override for the output directory.")
    return parser.parse_args()


def main() -> None:
    args = parse_args()
    config_path = Path(args.config)
    if not config_path.is_absolute():
        config_path = ROOT / config_path
    config = json.loads(config_path.read_text(encoding="utf-8"))

    output_dir = Path(args.output_dir) if args.output_dir else Path(config.get("output_dir", "outputs/daily_report"))
    if not output_dir.is_absolute():
        output_dir = ROOT / output_dir

    results = extract_sources(config, ROOT)
    workbook_path, summary_path, snapshot_path = build_output(results, config, output_dir, config_path)

    print(json.dumps({
        "report_name": config.get("report_name", "Daily Report"),
        "workbook": str(workbook_path),
        "summary": str(summary_path),
        "snapshot": str(snapshot_path),
        "sources_loaded": [result.name for result in results],
    }, indent=2))


if __name__ == "__main__":
    main()

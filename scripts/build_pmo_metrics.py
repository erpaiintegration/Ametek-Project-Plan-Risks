from __future__ import annotations

import argparse
import json
import re
import shutil
from dataclasses import dataclass
from datetime import datetime
from pathlib import Path
from typing import Any, Callable

import pandas as pd
from openpyxl.utils import get_column_letter

ROOT = Path(__file__).resolve().parents[1]


def canonical(value: Any) -> str:
    return " ".join(str(value or "").strip().lower().replace("_", " ").split())


def clean_text(value: Any) -> str:
    if value is None:
        return ""
    if pd.isna(value):
        return ""
    return str(value).strip()


def status_text(series: pd.Series) -> pd.Series:
    return series.fillna("").astype(str).str.strip().str.lower()


def as_numeric(series: pd.Series) -> pd.Series:
    cleaned = (
        series.astype(str)
        .str.replace(",", "", regex=False)
        .str.replace("%", "", regex=False)
        .str.replace("(", "-", regex=False)
        .str.replace(")", "", regex=False)
    )
    return pd.to_numeric(cleaned, errors="coerce")


def parse_workstream_descriptor(value: Any) -> tuple[str, str, str, str]:
    text = clean_text(value)
    if text in {"", "(Unassigned)"}:
        return "", "", "(Unassigned)", ""

    slash_parts = [part.strip() for part in text.split("/")]
    test_cycle = slash_parts[0] if slash_parts else ""
    location = slash_parts[1] if len(slash_parts) > 1 else ""
    workstream_and_test = "/".join(slash_parts[2:]).strip() if len(slash_parts) > 2 else ""

    if ":" in workstream_and_test:
        work_stream, descriptor_test_name = [part.strip() for part in workstream_and_test.split(":", 1)]
    else:
        work_stream, descriptor_test_name = workstream_and_test, ""

    return test_cycle, location, work_stream or "(Unassigned)", descriptor_test_name


@dataclass
class SourceFrame:
    source_name: str
    file_path: Path
    sheet_name: str
    dataframe: pd.DataFrame


@dataclass
class DomainBuildResult:
    tables: dict[str, pd.DataFrame]
    validation_rows: list[dict[str, Any]]
    validation_pq_order: list[str]


def build_testing_defects_validation_rows(domain_name: str) -> list[dict[str, Any]]:
    rows: list[dict[str, Any]] = []

    def add_pq_doc(
        pq_name: str,
        purpose: str,
        dependencies: str,
        notes: str,
    ) -> None:
        rows.append(
            {
                "Domain": domain_name,
                "RowType": "PQ",
                "Section": f"{domain_name} Metrics",
                "Item": pq_name,
                "MetricPQ": "",
                "FieldName": "",
                "MetricName": "",
                "Purpose": purpose,
                "Dependencies": dependencies,
                "Logic": notes,
            }
        )

    def add_metric_doc(
        metric_pq: str,
        field_name: str,
        metric_name: str,
        logic: str,
    ) -> None:
        rows.append(
            {
                "Domain": domain_name,
                "RowType": "Metric",
                "Section": f"{domain_name} Metrics",
                "Item": f"{metric_pq}.{field_name}",
                "MetricPQ": metric_pq,
                "FieldName": field_name,
                "MetricName": metric_name,
                "Purpose": "Validation metric field",
                "Dependencies": "PQ_TD_Detail" if metric_pq != "PQ_TD_Detail" else "PQ_TD_Source",
                "Logic": logic,
            }
        )

    add_pq_doc(
        pq_name="PQ_TD_Source",
        purpose="Raw source-of-truth ingest for Test/Defect workbook across included sheets.",
        dependencies="testing_defects_book source, include_sheet_patterns, header_row",
        notes="No business aggregation; preserves source lineage (file/sheet/row).",
    )
    add_pq_doc(
        pq_name="PQ_TD_Detail",
        purpose="Canonical row-level testing-defect detail with helper flags and section classification.",
        dependencies="PQ_TD_Source + alias mapping + section patterns",
        notes="Adds semantic flags and keys used by summary/workstream calculations and future milestone links.",
    )
    add_pq_doc(
        pq_name="PQ_TD_Summary",
        purpose="Aggregate totals for ALL and each Section (ITC/UAT/GO-LIVE).",
        dependencies="PQ_TD_Detail",
        notes="Group-by aggregation over Section with test/defect KPI calculations.",
    )
    add_pq_doc(
        pq_name="PQ_TD_Workstream",
        purpose="Workstream and owner drill-down metrics for dedicated Testing/Defects reporting.",
        dependencies="PQ_TD_Detail",
        notes="Includes both workstream totals and workstream+owner rows for resource-level reporting.",
    )

    add_metric_doc("PQ_TD_Detail", "Section", "Section bucket", "Regex-based classification from sheet/cycle/task/test text into ITC1/ITC2/ITC3/UAT/GO-LIVE/OTHER.")
    add_metric_doc("PQ_TD_Detail", "WorkstreamRaw", "Raw workstream descriptor", "Original composite descriptor string from source (before split).")
    add_metric_doc("PQ_TD_Detail", "TestCycle", "Parsed test cycle", "First segment before first '/' in WorkstreamRaw.")
    add_metric_doc("PQ_TD_Detail", "Location", "Parsed location", "Second segment between first and second '/' in WorkstreamRaw.")
    add_metric_doc("PQ_TD_Detail", "Workstream", "Parsed work stream", "Segment before ':' after cycle/location in WorkstreamRaw.")
    add_metric_doc("PQ_TD_Detail", "WorkstreamTestName", "Parsed descriptor test name", "Segment after ':' in WorkstreamRaw.")
    add_metric_doc("PQ_TD_Detail", "MilestoneLinkKey", "Milestone link key", "Uses TaskID when available, else TaskName for downstream milestone linkage.")
    add_metric_doc("PQ_TD_Detail", "HasDefectLink", "Defect linkage flag", "True when DefectCount > 0 OR DefectID populated OR (DefectsNew+DefectsInProgress+DefectsClosed) > 0.")
    add_metric_doc("PQ_TD_Detail", "IsNA", "Not Applicable flag", "Status contains 'not applicable' or 'n/a'.")
    add_metric_doc("PQ_TD_Detail", "IsPassed", "Passed flag", "Status contains 'pass'.")
    add_metric_doc("PQ_TD_Detail", "IsFailed", "Failed flag", "Status contains 'fail'.")
    add_metric_doc("PQ_TD_Detail", "IsBlocked", "Blocked flag", "Status contains 'block'.")
    add_metric_doc("PQ_TD_Detail", "IsInProgress", "In-progress flag", "Status contains in-progress style states and is not N/A.")
    add_metric_doc("PQ_TD_Detail", "IsNotRun", "Not-run flag", "Status contains not-run style states and is not N/A.")

    add_metric_doc("PQ_TD_Summary", "TotalTests", "Total tests", "COUNT rows in PQ_TD_Detail by Section (plus ALL rollup).")
    add_metric_doc("PQ_TD_Summary", "TestsNA", "Tests N/A", "SUM(IsNA).")
    add_metric_doc("PQ_TD_Summary", "ExecutableTests", "Executable tests", "TotalTests - TestsNA.")
    add_metric_doc("PQ_TD_Summary", "TestsPassed", "Tests passed", "SUM(IsPassed).")
    add_metric_doc("PQ_TD_Summary", "TestsFailed", "Tests failed", "SUM(IsFailed).")
    add_metric_doc("PQ_TD_Summary", "TestsBlocked", "Tests blocked", "SUM(IsBlocked).")
    add_metric_doc("PQ_TD_Summary", "TestsInProgress", "Tests in progress", "SUM(IsInProgress).")
    add_metric_doc("PQ_TD_Summary", "TestsNotRun", "Tests not run", "SUM(IsNotRun).")
    add_metric_doc("PQ_TD_Summary", "TestsWithDefectLink", "Tests with defects", "SUM(HasDefectLink).")
    add_metric_doc("PQ_TD_Summary", "DefectRefs", "Defect references", "SUM(DefectCount).")
    add_metric_doc("PQ_TD_Summary", "DefectsOpen", "Defects open/new", "SUM(DefectsNew).")
    add_metric_doc("PQ_TD_Summary", "DefectsInProgress", "Defects in progress", "SUM(DefectsInProgress).")
    add_metric_doc("PQ_TD_Summary", "DefectsClosed", "Defects closed", "SUM(DefectsClosed).")
    add_metric_doc("PQ_TD_Summary", "PassRatePct", "Pass rate %", "(TestsPassed / ExecutableTests) * 100; 0 when denominator is 0.")
    add_metric_doc("PQ_TD_Summary", "FailRatePct", "Fail rate %", "(TestsFailed / ExecutableTests) * 100; 0 when denominator is 0.")
    add_metric_doc("PQ_TD_Summary", "BlockedRatePct", "Blocked rate %", "(TestsBlocked / ExecutableTests) * 100; 0 when denominator is 0.")

    add_metric_doc("PQ_TD_Workstream", "TotalTests", "Total tests by workstream", "Same aggregation logic as summary but grouped by Section+Workstream (+Owner rows).")
    add_metric_doc("PQ_TD_Workstream", "PassRatePct", "Pass rate % by workstream", "(TestsPassed / ExecutableTests) * 100 within each grouping.")
    add_metric_doc("PQ_TD_Workstream", "DefectRefs", "Defect refs by workstream", "SUM(DefectCount) within each grouping.")
    add_metric_doc("PQ_TD_Workstream", "TestsWithDefectLink", "Tests with defects by workstream", "SUM(HasDefectLink) within each grouping.")

    return rows


def build_source_reference_table(loaded_sources: dict[str, list[SourceFrame]]) -> pd.DataFrame:
    rows: list[dict[str, Any]] = []
    for source_name, frames in loaded_sources.items():
        for frame in frames:
            try:
                last_modified = datetime.fromtimestamp(frame.file_path.stat().st_mtime).isoformat(timespec="seconds")
            except OSError:
                last_modified = ""
            rows.append(
                {
                    "SourceName": source_name,
                    "SourceFile": str(frame.file_path),
                    "SourceSheet": frame.sheet_name,
                    "RowCount": int(len(frame.dataframe)),
                    "ColumnCount": int(len(frame.dataframe.columns)),
                    "LastModified": last_modified,
                }
            )
    return pd.DataFrame(rows)


def build_metrics_validation_sheet(
    domains_cfg: list[dict[str, Any]],
    domain_results: dict[str, DomainBuildResult],
) -> pd.DataFrame:
    rows: list[dict[str, Any]] = []

    def add_row(
        domain: str,
        section: str,
        row_type: str,
        item: str = "",
        metric_pq: str = "",
        field_name: str = "",
        metric_name: str = "",
        purpose: str = "",
        dependencies: str = "",
        logic: str = "",
    ) -> None:
        rows.append(
            {
                "Domain": domain,
                "Section": section,
                "RowType": row_type,
                "Item": item,
                "MetricPQ": metric_pq,
                "FieldName": field_name,
                "MetricName": metric_name,
                "Purpose": purpose,
                "Dependencies": dependencies,
                "Logic": logic,
            }
        )

    for dcfg in domains_cfg:
        domain_name = str(dcfg.get("name", "domain"))
        section_title = str(dcfg.get("validation_section_title", f"{domain_name} Metrics"))
        domain_result = domain_results.get(domain_name)
        if domain_result is None:
            continue

        add_row(
            domain=domain_name,
            section=section_title,
            row_type="Section",
            item=section_title,
            purpose="Domain validation section",
        )

        validation_rows = domain_result.validation_rows or []
        pq_docs_cfg = (dcfg.get("validation") or {}).get("pq_docs") or {}
        metric_logic_cfg = (dcfg.get("validation") or {}).get("metric_logic") or {}
        expected_fields_cfg = (dcfg.get("validation") or {}).get("expected_fields") or {}

        pq_rows = [r for r in validation_rows if r.get("RowType") == "PQ"]
        pq_doc_map = {str(r.get("Item", "")): r for r in pq_rows}
        ordered_pqs = domain_result.validation_pq_order or sorted(domain_result.tables.keys())

        for pq_name in ordered_pqs:
            table_df = domain_result.tables.get(pq_name, pd.DataFrame())
            cfg_doc = pq_docs_cfg.get(pq_name, {})
            fallback_doc = pq_doc_map.get(pq_name, {})

            purpose = str(cfg_doc.get("purpose") or fallback_doc.get("Purpose") or f"PQ table for {pq_name}.")
            dependencies = str(cfg_doc.get("dependencies") or fallback_doc.get("Dependencies") or "")
            logic = str(cfg_doc.get("notes") or fallback_doc.get("Logic") or "")

            add_row(
                domain=domain_name,
                section=section_title,
                row_type="PQ",
                item=pq_name,
                purpose=purpose,
                dependencies=dependencies,
                logic=logic,
            )

            add_row(
                domain=domain_name,
                section=section_title,
                row_type="MetricGroup",
                item=f"{pq_name} fields",
                metric_pq=pq_name,
                purpose="Field-level metrics/attributes for validation",
            )

            table_cols = [str(c) for c in table_df.columns.tolist()]
            expected_cols = [str(c) for c in expected_fields_cfg.get(pq_name, [])]
            metric_logic_cols = [str(c) for c in (metric_logic_cfg.get(pq_name, {}) or {}).keys()]

            cols_to_document = list(dict.fromkeys(table_cols + expected_cols + metric_logic_cols))
            for col in cols_to_document:
                logic_text = str(metric_logic_cfg.get(pq_name, {}).get(col, "Direct field from PQ output for validation traceability."))
                add_row(
                    domain=domain_name,
                    section=section_title,
                    row_type="Metric",
                    item=f"{pq_name}.{col}",
                    metric_pq=pq_name,
                    field_name=str(col),
                    metric_name=str(col),
                    purpose="Validation field",
                    dependencies=pq_name,
                    logic=logic_text,
                )

        add_row(
            domain=domain_name,
            section=section_title,
            row_type="Spacer",
            item="",
        )

    return pd.DataFrame(rows)


def resolve_source_path(source: dict[str, Any]) -> Path:
    if source.get("path"):
        path = Path(source["path"])
        return path if path.is_absolute() else ROOT / path

    if source.get("path_glob"):
        glob_str = source["path_glob"]
        glob_path = Path(glob_str)
        base: Path
        if glob_path.is_absolute():
            parts = glob_path.parts
            base_parts: list[str] = []
            pattern_parts: list[str] = []
            found_wildcard = False
            for part in parts:
                if not found_wildcard and ("*" not in part and "?" not in part and "[" not in part):
                    base_parts.append(part)
                else:
                    found_wildcard = True
                    pattern_parts.append(part)
            if not pattern_parts:
                candidate = Path(*parts)
                if not candidate.exists():
                    raise FileNotFoundError(f"File not found: {candidate}")
                return candidate
            base = Path(*base_parts) if len(base_parts) > 1 else Path(base_parts[0])
            pattern = "/".join(pattern_parts)
            matches = list(base.glob(pattern))
        else:
            base = ROOT
            matches = list(ROOT.glob(glob_str))

        if not matches:
            raise FileNotFoundError(f"No files matched glob '{glob_str}'")

        return _select_glob_match(source=source, matches=matches, base_dir=base)

    raise ValueError(f"Source '{source.get('name', '<unnamed>')}' must define either 'path' or 'path_glob'.")


def _is_archive_path(path: Path) -> bool:
    return any("archive" in part.lower() for part in path.parts)


def _select_glob_match(source: dict[str, Any], matches: list[Path], base_dir: Path) -> Path:
    source_name = str(source.get("name", "source"))
    pick = str(source.get("pick", "latest")).lower()

    use_file_hygiene = bool(source.get("auto_manage_duplicates", False))
    if not use_file_hygiene:
        if pick == "latest":
            matches.sort(key=lambda p: p.stat().st_mtime, reverse=True)
            return matches[0]
        matches.sort()
        return matches[0]

    active_matches = [p for p in matches if not _is_archive_path(p)]
    archived_matches = [p for p in matches if _is_archive_path(p)]

    restore_from_archive = bool(source.get("restore_latest_from_archive", True))
    if not active_matches and restore_from_archive and archived_matches:
        archived_matches.sort(key=lambda p: p.stat().st_mtime, reverse=True)
        archived_latest = archived_matches[0]

        restore_dir = Path(str(source.get("restore_dir", base_dir)))
        restore_dir.mkdir(parents=True, exist_ok=True)
        restored_target = restore_dir / archived_latest.name

        if archived_latest != restored_target:
            if restored_target.exists():
                existing_mtime = restored_target.stat().st_mtime
                archived_mtime = archived_latest.stat().st_mtime
                if archived_mtime > existing_mtime:
                    restored_target.unlink()
                    shutil.move(str(archived_latest), str(restored_target))
                else:
                    # Keep existing active file and leave archived copy in place
                    pass
            else:
                shutil.move(str(archived_latest), str(restored_target))

        if restored_target.exists():
            print(f"[info] Restored latest archived source for '{source_name}' to: {restored_target}")

        active_matches = [p for p in matches if p.exists() and not _is_archive_path(p)]
        if restored_target.exists() and restored_target not in active_matches:
            active_matches.append(restored_target)

    candidate_pool = active_matches if active_matches else [p for p in matches if p.exists()]
    if not candidate_pool:
        raise FileNotFoundError(f"No usable files remained for source '{source_name}' after duplicate management.")

    if pick == "latest":
        candidate_pool.sort(key=lambda p: p.stat().st_mtime, reverse=True)
    else:
        candidate_pool.sort()
    selected = candidate_pool[0]

    archive_older = bool(source.get("archive_older_matches", True))
    if archive_older:
        archive_folder_name = str(source.get("duplicates_archive_folder", "archive_dedup"))
        archive_dir = base_dir / archive_folder_name
        archive_dir.mkdir(parents=True, exist_ok=True)

        older_candidates = [p for p in candidate_pool if p != selected and p.exists() and not _is_archive_path(p)]
        for older in older_candidates:
            target = archive_dir / older.name
            if target.exists():
                timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
                target = archive_dir / f"{older.stem}_{timestamp}{older.suffix}"
            shutil.move(str(older), str(target))
            print(f"[info] Archived older duplicate for '{source_name}': {older} -> {target}")

    return selected


def _sheet_allowed(sheet_name: str, source: dict[str, Any]) -> bool:
    include_patterns = source.get("include_sheet_patterns") or []
    exclude_patterns = source.get("exclude_sheet_patterns") or []

    if include_patterns:
        if not any(re.search(pattern, sheet_name, flags=re.IGNORECASE) for pattern in include_patterns):
            return False

    if exclude_patterns:
        if any(re.search(pattern, sheet_name, flags=re.IGNORECASE) for pattern in exclude_patterns):
            return False

    return True


def load_source_frames(source: dict[str, Any]) -> list[SourceFrame]:
    source_name = source["name"]
    required = bool(source.get("required", True))
    header_row = int(source.get("header_row", 0))
    read_all_sheets = bool(source.get("read_all_sheets", False))

    try:
        source_path = resolve_source_path(source)
    except FileNotFoundError as exc:
        if required:
            raise
        print(f"[warn] Optional source skipped: {exc}")
        return []

    ext = source_path.suffix.lower()
    frames: list[SourceFrame] = []

    if ext == ".csv":
        df = pd.read_csv(source_path, header=header_row)
        frames.append(SourceFrame(source_name=source_name, file_path=source_path, sheet_name="csv", dataframe=df))
        return frames

    if ext not in {".xlsx", ".xlsm", ".xls"}:
        raise ValueError(f"Unsupported source file type: {source_path.suffix}")

    if read_all_sheets:
        workbook_frames = pd.read_excel(source_path, sheet_name=None, header=header_row)
        for sheet_name, df in workbook_frames.items():
            if not _sheet_allowed(str(sheet_name), source):
                continue
            frames.append(
                SourceFrame(
                    source_name=source_name,
                    file_path=source_path,
                    sheet_name=str(sheet_name),
                    dataframe=df,
                )
            )
    else:
        sheet_name = source.get("sheet_name", 0)
        df = pd.read_excel(source_path, sheet_name=sheet_name, header=header_row)
        frames.append(
            SourceFrame(
                source_name=source_name,
                file_path=source_path,
                sheet_name=str(sheet_name),
                dataframe=df,
            )
        )

    if not frames and required:
        raise ValueError(f"Source '{source_name}' loaded but no sheets matched include/exclude filters.")

    return frames


def _build_lookup(columns: list[str]) -> dict[str, str]:
    return {canonical(col): col for col in columns}


def _pick_column(lookup: dict[str, str], candidates: list[str]) -> str | None:
    for candidate in candidates:
        hit = lookup.get(canonical(candidate))
        if hit:
            return hit
    return None


def _extract_series(df: pd.DataFrame, aliases: dict[str, list[str]], key: str) -> pd.Series:
    lookup = _build_lookup([str(c) for c in df.columns])
    col = _pick_column(lookup, aliases.get(key, []))
    if col is None:
        return pd.Series([pd.NA] * len(df), index=df.index)
    return df[col]


def _classify_section(text: str, section_patterns: dict[str, list[str]], fallback: str = "OTHER") -> str:
    for section, patterns in section_patterns.items():
        for pattern in patterns:
            if re.search(pattern, text, flags=re.IGNORECASE):
                return section
    return fallback


def build_testing_defects_domain(domain_config: dict[str, Any], source_frames: dict[str, list[SourceFrame]]) -> DomainBuildResult:
    source_names = domain_config.get("source_names", [])
    aliases: dict[str, list[str]] = domain_config.get("aliases", {})
    section_patterns: dict[str, list[str]] = domain_config.get("section_patterns", {
        "ITC1": [r"\bitc\s*1\b", r"\bitc1\b"],
        "ITC2": [r"\bitc\s*2\b", r"\bitc2\b"],
        "ITC3": [r"\bitc\s*3\b", r"\bitc3\b"],
        "UAT": [r"\buat\b", r"user\s*accept"],
        "GO-LIVE": [r"go\s*-?\s*live", r"hypercare", r"cutover"],
    })

    records: list[pd.DataFrame] = []
    domain_name = str(domain_config.get("name", "testing_defects"))
    validation_rows = build_testing_defects_validation_rows(domain_name)
    for source_name in source_names:
        for frame in source_frames.get(source_name, []):
            raw = frame.dataframe.copy()

            raw.insert(0, "source_row_number", raw.index + 1)
            raw.insert(0, "source_sheet", frame.sheet_name)
            raw.insert(0, "source_file", str(frame.file_path))
            raw.insert(0, "source_name", frame.source_name)
            records.append(raw)

    if not records:
        empty = pd.DataFrame()
        return DomainBuildResult(
            tables={
                "PQ_TD_Source": empty,
                "PQ_TD_Detail": empty,
                "PQ_TD_Summary": empty,
                "PQ_TD_Workstream": empty,
            },
            validation_rows=validation_rows,
            validation_pq_order=[
                "PQ_TD_Source",
                "PQ_TD_Detail",
                "PQ_TD_Summary",
                "PQ_TD_Workstream",
            ],
        )

    source_df = pd.concat(records, ignore_index=True)

    test_id = _extract_series(source_df, aliases, "test_id").map(clean_text)
    test_name = _extract_series(source_df, aliases, "test_name").map(clean_text)
    status = _extract_series(source_df, aliases, "status")
    workstream_raw = _extract_series(source_df, aliases, "workstream").map(clean_text)
    parsed_workstream = workstream_raw.map(parse_workstream_descriptor)
    parsed_workstream_df = pd.DataFrame(
        parsed_workstream.tolist(),
        columns=["TestCycle", "Location", "Workstream", "WorkstreamTestName"],
        index=source_df.index,
    )
    owner = _extract_series(source_df, aliases, "owner").map(clean_text)
    cycle = _extract_series(source_df, aliases, "cycle").map(clean_text)
    task_id = _extract_series(source_df, aliases, "task_id").map(clean_text)
    task_name = _extract_series(source_df, aliases, "task_name").map(clean_text)
    planned_date = pd.to_datetime(_extract_series(source_df, aliases, "planned_date"), errors="coerce")
    executed_date = pd.to_datetime(_extract_series(source_df, aliases, "executed_date"), errors="coerce")

    defect_id = _extract_series(source_df, aliases, "defect_id").map(clean_text)
    defect_count = as_numeric(_extract_series(source_df, aliases, "defect_count")).fillna(0)
    defects_new = as_numeric(_extract_series(source_df, aliases, "defects_new")).fillna(0)
    defects_in_progress = as_numeric(_extract_series(source_df, aliases, "defects_in_progress")).fillna(0)
    defects_closed = as_numeric(_extract_series(source_df, aliases, "defects_closed")).fillna(0)

    status_norm = status_text(status)
    is_na = status_norm.str.contains(r"not\s*applicable|\bn/a\b", regex=True)
    is_passed = status_norm.str.contains(r"\bpass", regex=True)
    is_failed = status_norm.str.contains(r"\bfail", regex=True)
    is_blocked = status_norm.str.contains(r"\bblock", regex=True)
    is_in_progress = status_norm.str.contains(r"in\s*progress|in-progress|inprogress|pending|queued", regex=True) & (~is_na)
    is_not_run = status_norm.str.contains(r"not\s*run|not\s*executed|todo|to\s*do", regex=True) & (~is_na)

    has_defect_link = (defect_count > 0) | (defect_id != "") | ((defects_new + defects_in_progress + defects_closed) > 0)

    section_text = (
        source_df["source_sheet"].map(clean_text)
        + " | "
        + cycle
        + " | "
        + task_name
        + " | "
        + test_name
    )
    section = section_text.map(lambda txt: _classify_section(txt, section_patterns))

    detail_df = pd.DataFrame(
        {
            "Section": section,
            "WorkstreamRaw": workstream_raw.replace("", "(Unassigned)"),
            "TestCycle": parsed_workstream_df["TestCycle"],
            "Location": parsed_workstream_df["Location"],
            "Workstream": parsed_workstream_df["Workstream"],
            "WorkstreamTestName": parsed_workstream_df["WorkstreamTestName"],
            "Owner": owner.replace("", "(Unassigned)"),
            "SourceSheet": source_df["source_sheet"].map(clean_text),
            "SourceFile": source_df["source_file"].map(clean_text),
            "SourceRow": source_df["source_row_number"],
            "TestID": test_id,
            "TestName": test_name,
            "Status": status.map(clean_text),
            "Cycle": cycle,
            "PlannedDate": planned_date,
            "ExecutedDate": executed_date,
            "TaskID": task_id,
            "TaskName": task_name,
            "MilestoneLinkKey": task_id.where(task_id != "", task_name),
            "DefectID": defect_id,
            "DefectCount": defect_count,
            "DefectsNew": defects_new,
            "DefectsInProgress": defects_in_progress,
            "DefectsClosed": defects_closed,
            "HasDefectLink": has_defect_link,
            "IsNA": is_na,
            "IsPassed": is_passed,
            "IsFailed": is_failed,
            "IsBlocked": is_blocked,
            "IsInProgress": is_in_progress,
            "IsNotRun": is_not_run,
        }
    )

    # Final guardrail: remove rows that remain effectively blank after canonical mapping.
    # Also exclude status-only artifact rows that have no test identity.
    if not detail_df.empty:
        na_excluded_mask = ~detail_df["IsNA"].fillna(False)
        dropped_na_rows = int((~na_excluded_mask).sum())
        detail_df = detail_df.loc[na_excluded_mask].copy()

        has_test_identity = detail_df["TestID"].map(clean_text).ne("") | detail_df["TestName"].map(clean_text).ne("")
        has_defect_identity = (
            detail_df["DefectID"].map(clean_text).ne("")
            | detail_df["DefectCount"].fillna(0).gt(0)
            | detail_df["DefectsNew"].fillna(0).gt(0)
            | detail_df["DefectsInProgress"].fillna(0).gt(0)
            | detail_df["DefectsClosed"].fillna(0).gt(0)
        )

        meaningful_mask = (
            has_test_identity
            | has_defect_identity
        )
        dropped_blank_detail_rows = int((~meaningful_mask).sum())
        detail_df = detail_df.loc[meaningful_mask].copy()
    else:
        dropped_na_rows = 0
        dropped_blank_detail_rows = 0

    if dropped_na_rows or dropped_blank_detail_rows:
        print(
            f"[info] testing_defects blank-row cleanup: "
            f"na_status_dropped={dropped_na_rows}, "
            f"detail_dropped={dropped_blank_detail_rows}"
        )

    def _aggregate(group_cols: list[str]) -> pd.DataFrame:
        agg_spec = {
            "TotalTests": ("TestID", "count"),
            "TestsNA": ("IsNA", "sum"),
            "TestsPassed": ("IsPassed", "sum"),
            "TestsFailed": ("IsFailed", "sum"),
            "TestsBlocked": ("IsBlocked", "sum"),
            "TestsInProgress": ("IsInProgress", "sum"),
            "TestsNotRun": ("IsNotRun", "sum"),
            "TestsWithDefectLink": ("HasDefectLink", "sum"),
            "DefectRefs": ("DefectCount", "sum"),
            "DefectsOpen": ("DefectsNew", "sum"),
            "DefectsInProgress": ("DefectsInProgress", "sum"),
            "DefectsClosed": ("DefectsClosed", "sum"),
        }

        if group_cols:
            grouped = detail_df.groupby(group_cols, dropna=False).agg(**agg_spec).reset_index()
        else:
            grouped = pd.DataFrame(
                [
                    {
                        "TotalTests": int(detail_df["TestID"].count()),
                        "TestsNA": int(detail_df["IsNA"].sum()),
                        "TestsPassed": int(detail_df["IsPassed"].sum()),
                        "TestsFailed": int(detail_df["IsFailed"].sum()),
                        "TestsBlocked": int(detail_df["IsBlocked"].sum()),
                        "TestsInProgress": int(detail_df["IsInProgress"].sum()),
                        "TestsNotRun": int(detail_df["IsNotRun"].sum()),
                        "TestsWithDefectLink": int(detail_df["HasDefectLink"].sum()),
                        "DefectRefs": float(detail_df["DefectCount"].sum()),
                        "DefectsOpen": float(detail_df["DefectsNew"].sum()),
                        "DefectsInProgress": float(detail_df["DefectsInProgress"].sum()),
                        "DefectsClosed": float(detail_df["DefectsClosed"].sum()),
                    }
                ]
            )

        grouped["ExecutableTests"] = grouped["TotalTests"] - grouped["TestsNA"]
        grouped["PassRatePct"] = (grouped["TestsPassed"] / grouped["ExecutableTests"].replace(0, pd.NA) * 100).fillna(0).round(1)
        grouped["FailRatePct"] = (grouped["TestsFailed"] / grouped["ExecutableTests"].replace(0, pd.NA) * 100).fillna(0).round(1)
        grouped["BlockedRatePct"] = (grouped["TestsBlocked"] / grouped["ExecutableTests"].replace(0, pd.NA) * 100).fillna(0).round(1)
        return grouped

    summary_by_section = _aggregate(["Section"]) 
    summary_all = _aggregate([])
    if not summary_all.empty:
        summary_all.insert(0, "Section", "ALL")
    summary_df = pd.concat([summary_all, summary_by_section], ignore_index=True)

    ws_rollup = _aggregate(["Section", "Workstream"])
    ws_rollup.insert(2, "Owner", "(All Owners)")
    ws_owner = _aggregate(["Section", "Workstream", "Owner"])
    workstream_df = pd.concat([ws_rollup, ws_owner], ignore_index=True)

    detail_df = detail_df.sort_values(["Section", "Workstream", "Owner", "TestID", "SourceRow"], kind="stable")
    summary_df = summary_df.sort_values(["Section"], kind="stable")
    workstream_df = workstream_df.sort_values(["Section", "Workstream", "Owner"], kind="stable")

    return DomainBuildResult(
        tables={
            "PQ_TD_Source": source_df,
            "PQ_TD_Detail": detail_df,
            "PQ_TD_Summary": summary_df,
            "PQ_TD_Workstream": workstream_df,
        },
        validation_rows=validation_rows,
        validation_pq_order=[
            "PQ_TD_Source",
            "PQ_TD_Detail",
            "PQ_TD_Summary",
            "PQ_TD_Workstream",
        ],
    )


DOMAIN_BUILDERS: dict[str, Callable[[dict[str, Any], dict[str, list[SourceFrame]]], DomainBuildResult]] = {
    "testing_defects": build_testing_defects_domain,
}


def prepare_sheet(df: pd.DataFrame) -> pd.DataFrame:
    if df.empty:
        return pd.DataFrame({"info": ["No data available"]})
    prepared = df.copy()
    for col in prepared.columns:
        prepared[col] = prepared[col].map(
            lambda v: v.isoformat() if isinstance(v, (pd.Timestamp, datetime)) and not pd.isna(v) else v
        )
    return prepared


def _stringify(value: Any) -> str:
    if value is None:
        return ""
    if isinstance(value, (dict, list, tuple, set)):
        return json.dumps(value, ensure_ascii=False)
    return str(value)


def build_configuration_sheet(config: dict[str, Any], domains_cfg: list[dict[str, Any]], source_catalog: dict[str, dict[str, Any]]) -> pd.DataFrame:
    rows: list[dict[str, Any]] = []

    def add_row(scope: str, item: str, parameter: str, value: Any, default: Any, description: str) -> None:
        rows.append(
            {
                "Scope": scope,
                "Item": item,
                "Parameter": parameter,
                "Value": _stringify(value),
                "Default": _stringify(default),
                "Description": description,
            }
        )

    add_row(
        "Global",
        "runtime",
        "contract_mode",
        config.get("contract_mode", "hybrid"),
        "hybrid",
        "Governance mode: metrics owns KPI definitions; display can do presentation shaping only.",
    )
    add_row(
        "Global",
        "runtime",
        "metrics_workbook_name",
        config.get("metrics_workbook_name", "Ametek PMO Metrics.xlsx"),
        "Ametek PMO Metrics.xlsx",
        "Output workbook file name.",
    )
    add_row(
        "Global",
        "runtime",
        "output_dir",
        config.get("output_dir", "outputs/pmo_metrics"),
        "outputs/pmo_metrics",
        "Output folder for workbook and table CSV exports.",
    )
    add_row(
        "Global",
        "runtime",
        "export_tables_csv",
        config.get("export_tables_csv", False),
        False,
        "When true, also writes each sheet table to /tables as CSV; workbook is always generated.",
    )

    source_defaults: dict[str, Any] = {
        "path": "",
        "path_glob": "",
        "pick": "latest",
        "required": True,
        "header_row": 0,
        "read_all_sheets": False,
        "include_sheet_patterns": [],
        "exclude_sheet_patterns": [],
        "auto_manage_duplicates": False,
        "restore_latest_from_archive": True,
        "restore_dir": "",
        "archive_older_matches": True,
        "duplicates_archive_folder": "archive_dedup",
    }

    source_descriptions: dict[str, str] = {
        "path": "Absolute/relative source path (single file mode).",
        "path_glob": "Glob pattern for source discovery (supports recursive **).",
        "pick": "Match selection strategy (latest or sorted-first).",
        "required": "Fail run when source is missing if true.",
        "header_row": "Zero-based header row index.",
        "read_all_sheets": "When true, ingests all included sheets from workbook.",
        "include_sheet_patterns": "Regex patterns of sheet names to include.",
        "exclude_sheet_patterns": "Regex patterns of sheet names to exclude.",
        "auto_manage_duplicates": "Enable newest-file selection and duplicate archiving logic.",
        "restore_latest_from_archive": "Restore newest archived file when no active file found.",
        "restore_dir": "Target folder where restored file is placed.",
        "archive_older_matches": "Archive older duplicate files after newest selection.",
        "duplicates_archive_folder": "Archive subfolder name for deduplicated files.",
    }

    for source_name, src in source_catalog.items():
        for param, default in source_defaults.items():
            add_row(
                "Source",
                source_name,
                param,
                src.get(param, default),
                default,
                source_descriptions.get(param, ""),
            )

    domain_defaults: dict[str, Any] = {
        "enabled": True,
        "builder": "",
        "validation_section_title": "",
        "source_names": [],
        "section_patterns": {},
        "aliases": {},
        "validation": {},
    }
    domain_descriptions: dict[str, str] = {
        "enabled": "Enable/disable this domain for the run.",
        "builder": "Domain builder function key used by the framework.",
        "validation_section_title": "Section title shown in Metrics Validation sheet.",
        "source_names": "List of source_catalog names consumed by domain.",
        "section_patterns": "Regex classification map for logical sections (ITC/UAT/GO-LIVE).",
        "aliases": "Column alias mapping from source headers to canonical fields.",
        "validation": "Optional PQ docs and field logic text used in validation output.",
    }

    for domain in domains_cfg:
        domain_name = str(domain.get("name", "domain"))
        for param, default in domain_defaults.items():
            add_row(
                "Domain",
                domain_name,
                param,
                domain.get(param, default),
                default,
                domain_descriptions.get(param, ""),
            )

    add_row(
        "HowTo",
        "editing",
        "where_to_change",
        "Update config/pmo-metrics-sources.example.json (or your runtime config copy), then rerun the script.",
        "",
        "The Configuration sheet is a live snapshot/reference; JSON is the source of truth.",
    )

    return pd.DataFrame(rows)


def build_contract_sheet(config: dict[str, Any]) -> pd.DataFrame:
    mode = str(config.get("contract_mode", "hybrid")).lower().strip() or "hybrid"
    if mode != "hybrid":
        mode = "hybrid"

    rows: list[dict[str, Any]] = [
        {
            "Contract": "PMO Metrics ↔ Huddle Display",
            "Mode": mode,
            "Rule": "Boundary",
            "Definition": "Metrics workbook is source of truth for business metric definitions; display workbook is presentation layer.",
        },
        {
            "Contract": "PMO Metrics ↔ Huddle Display",
            "Mode": mode,
            "Rule": "Metrics responsibilities",
            "Definition": "Ingest sources, normalize schema, compute reusable KPIs, publish canonical domain PQ tables (detail/summary/workstream).",
        },
        {
            "Contract": "PMO Metrics ↔ Huddle Display",
            "Mode": mode,
            "Rule": "Display responsibilities",
            "Definition": "Consume canonical metrics PQ tables and perform presentation shaping only (sorting, grouping, top-N, chart helper bins, layout flow).",
        },
        {
            "Contract": "PMO Metrics ↔ Huddle Display",
            "Mode": mode,
            "Rule": "Not allowed in display",
            "Definition": "Do not redefine KPI logic/thresholds/semantic flags in huddle workbook or display builder.",
        },
        {
            "Contract": "PMO Metrics ↔ Huddle Display",
            "Mode": mode,
            "Rule": "Canonical testing tables",
            "Definition": "Display builder should read PQ_TD_Detail, PQ_TD_Summary, PQ_TD_Workstream (and other canonical domain PQs as added).",
        },
        {
            "Contract": "PMO Metrics ↔ Huddle Display",
            "Mode": mode,
            "Rule": "Change management",
            "Definition": "Any KPI definition change must be implemented in metrics builder and documented in Metrics Validation sheet.",
        },
        {
            "Contract": "PMO Metrics ↔ Huddle Display",
            "Mode": mode,
            "Rule": "LastUpdated",
            "Definition": datetime.now().isoformat(timespec="seconds"),
        },
    ]
    return pd.DataFrame(rows)


def build_visual_design_contract_sheet() -> pd.DataFrame:
    rows: list[dict[str, Any]] = [
        {
            "Section": "Purpose",
            "Category": "Control Tower",
            "Rule": "Metrics workbook is the UI governance source of truth",
            "Before": "Dashboard styling standards scattered across scripts/chats",
            "After": "Single reference checklist in metrics workbook",
            "PassCriteria": "Sheet exists and is referenced by all dashboard domain builders",
            "Notes": "Use with docs/PMO_DASHBOARD_DESIGN_CONTRACT.md",
        },
        {
            "Section": "Layout",
            "Category": "3-Zone Structure",
            "Rule": "Left KPI cards, center tables, right charts",
            "Before": "Inconsistent placement and reading flow",
            "After": "Consistent information architecture across domains",
            "PassCriteria": "All domain dashboards follow same panel zones",
            "Notes": "Supports faster executive scan",
        },
        {
            "Section": "Styling",
            "Category": "Tokenized Design",
            "Rule": "Use centralized color/font/border tokens",
            "Before": "Hardcoded ad-hoc values causing visual drift",
            "After": "Brand-consistent and maintainable design system",
            "PassCriteria": "No random style literals in dashboard scripts",
            "Notes": "Use shared token constants",
        },
        {
            "Section": "Charts",
            "Category": "Readability",
            "Rule": "Top-N categories, shortened labels, subtitle context",
            "Before": "Cluttered axis labels and noisy interpretation",
            "After": "Clear chart meaning at first glance",
            "PassCriteria": "Chart subtitles present and labels readable at 100% zoom",
            "Notes": "Use top 10 and label shortening",
        },
        {
            "Section": "Governance",
            "Category": "Metrics vs Display",
            "Rule": "KPI logic remains in metrics layer only",
            "Before": "Presentation layer sometimes redefined semantics",
            "After": "Display layer does shaping only",
            "PassCriteria": "No KPI semantic overrides in huddle scripts",
            "Notes": "Hybrid contract enforcement",
        },
        {
            "Section": "Release",
            "Category": "Definition of Done",
            "Rule": "Before/after checklist must pass prior to rollout",
            "Before": "Subjective review outcomes",
            "After": "Objective release gate",
            "PassCriteria": "All required rows marked pass or approved exception",
            "Notes": "Attach screenshot evidence when needed",
        },
    ]
    return pd.DataFrame(rows)


def write_sheet(writer: pd.ExcelWriter, sheet_name: str, df: pd.DataFrame) -> None:
    safe_sheet = sheet_name[:31]
    prepared = prepare_sheet(df)
    prepared.to_excel(writer, sheet_name=safe_sheet, index=False)
    ws = writer.sheets[safe_sheet]
    ws.freeze_panes = "A2"
    for idx, col in enumerate(prepared.columns, start=1):
        values = [clean_text(col)] + [clean_text(v) for v in prepared[col].head(200)]
        width = min(max(len(v) for v in values) + 2, 60)
        ws.column_dimensions[get_column_letter(idx)].width = max(width, 12)
    ws.auto_filter.ref = ws.dimensions


def parse_args() -> argparse.Namespace:
    parser = argparse.ArgumentParser(
        description="Build modular PMO metrics workbook (domain-pluggable)."
    )
    parser.add_argument("--config", required=True, help="Path to PMO metrics config JSON.")
    parser.add_argument("--output-dir", help="Optional output directory override.")
    parser.add_argument(
        "--domains",
        nargs="*",
        help="Optional list of domain names to run. If omitted, runs enabled domains from config.",
    )
    return parser.parse_args()


def main() -> None:
    args = parse_args()

    config_path = Path(args.config)
    if not config_path.is_absolute():
        config_path = ROOT / config_path
    config = json.loads(config_path.read_text(encoding="utf-8"))

    configured_output = config.get("output_dir", "outputs/pmo_metrics")
    output_dir = Path(args.output_dir) if args.output_dir else Path(configured_output)
    if not output_dir.is_absolute():
        output_dir = ROOT / output_dir
    output_dir.mkdir(parents=True, exist_ok=True)

    source_catalog = {src["name"]: src for src in config.get("source_catalog", [])}
    domains_cfg = config.get("domains", [])

    if args.domains:
        requested = {d.strip() for d in args.domains if d.strip()}
        domains_cfg = [d for d in domains_cfg if d.get("name") in requested]
    else:
        domains_cfg = [d for d in domains_cfg if d.get("enabled", True)]

    if not domains_cfg:
        raise ValueError("No domains selected. Check config or --domains argument.")

    needed_source_names: set[str] = set()
    for dcfg in domains_cfg:
        needed_source_names.update(dcfg.get("source_names", []))

    loaded_sources: dict[str, list[SourceFrame]] = {}
    for source_name in sorted(needed_source_names):
        if source_name not in source_catalog:
            raise KeyError(f"Domain references unknown source '{source_name}'.")
        loaded_sources[source_name] = load_source_frames(source_catalog[source_name])

    all_tables: dict[str, pd.DataFrame] = {}
    metadata_rows: list[dict[str, Any]] = []
    domain_results: dict[str, DomainBuildResult] = {}
    for dcfg in domains_cfg:
        builder_name = dcfg.get("builder")
        domain_name = dcfg.get("name")
        if builder_name not in DOMAIN_BUILDERS:
            raise KeyError(f"Unknown domain builder '{builder_name}' for domain '{domain_name}'.")

        domain_result = DOMAIN_BUILDERS[builder_name](dcfg, loaded_sources)
        domain_results[str(domain_name)] = domain_result
        all_tables.update(domain_result.tables)

        metadata_rows.append(
            {
                "Domain": domain_name,
                "Builder": builder_name,
                "SourceNames": ", ".join(dcfg.get("source_names", [])),
                "TablesProduced": ", ".join(sorted(domain_result.tables.keys())),
                "GeneratedAt": datetime.now().isoformat(timespec="seconds"),
            }
        )

    metadata_df = pd.DataFrame(metadata_rows)
    validation_df = build_metrics_validation_sheet(domains_cfg=domains_cfg, domain_results=domain_results)
    source_ref_df = build_source_reference_table(loaded_sources)

    all_tables["PQ_Source_Reference"] = source_ref_df
    all_tables["PQ_Run_Metadata"] = metadata_df
    all_tables["Metrics Validation"] = validation_df
    all_tables["Contract"] = build_contract_sheet(config)
    all_tables["visual design Contract"] = build_visual_design_contract_sheet()
    all_tables["Configuration"] = build_configuration_sheet(
        config=config,
        domains_cfg=domains_cfg,
        source_catalog=source_catalog,
    )

    workbook_name = str(config.get("metrics_workbook_name", "Ametek PMO Metrics.xlsx"))
    workbook_path = output_dir / workbook_name
    with pd.ExcelWriter(workbook_path, engine="openpyxl") as writer:
        for sheet_name, df in all_tables.items():
            write_sheet(writer, sheet_name, df)

    export_tables_csv = bool(config.get("export_tables_csv", False))
    tables_dir = output_dir / "tables"
    if export_tables_csv:
        tables_dir.mkdir(parents=True, exist_ok=True)
        for table_name, df in all_tables.items():
            df.to_csv(tables_dir / f"{table_name}.csv", index=False, encoding="utf-8")

    print(
        json.dumps(
            {
                "workbook": str(workbook_path),
                "tables_dir": str(tables_dir) if export_tables_csv else "",
                "export_tables_csv": export_tables_csv,
                "domains": [d.get("name") for d in domains_cfg],
                "tables": sorted(all_tables.keys()),
            },
            indent=2,
        )
    )


if __name__ == "__main__":
    main()

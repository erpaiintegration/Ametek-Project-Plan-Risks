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


def build_raid_validation_rows(domain_name: str) -> list[dict[str, Any]]:
    rows: list[dict[str, Any]] = []

    def add_pq_doc(pq_name: str, purpose: str, dependencies: str, notes: str) -> None:
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

    def add_metric_doc(metric_pq: str, field_name: str, metric_name: str, logic: str) -> None:
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
                "Dependencies": "PQ_RAID_Detail" if metric_pq != "PQ_RAID_Detail" else "PQ_RAID_Source",
                "Logic": logic,
            }
        )

    add_pq_doc(
        pq_name="PQ_RAID_Source",
        purpose="Raw source-of-truth ingest for RAID workbook across included sheets.",
        dependencies="raid_log_book source, include_sheet_patterns, header_row",
        notes="No aggregation; preserves source lineage (file/sheet/row).",
    )
    add_pq_doc(
        pq_name="PQ_RAID_Detail",
        purpose="Canonical RAID row-level detail with type/status/priority helper flags.",
        dependencies="PQ_RAID_Source + alias mapping + RAID type inference",
        notes="Adds IsOpen/IsClosed/IsOverdue/IsHighPriority and due-window helper fields.",
    )
    add_pq_doc(
        pq_name="PQ_RAID_Summary",
        purpose="Aggregate totals for ALL and each RAID type (Risk/Issue/Action/etc.).",
        dependencies="PQ_RAID_Detail",
        notes="Group-by aggregation over RaidType with open/closed/overdue/high-priority KPIs.",
    )
    add_pq_doc(
        pq_name="PQ_RAID_Workstream",
        purpose="RAID drill-down metrics by RAID type, workstream, and owner.",
        dependencies="PQ_RAID_Detail",
        notes="Includes both workstream totals and workstream+owner rows for accountability reporting.",
    )

    add_metric_doc("PQ_RAID_Detail", "RaidType", "RAID type bucket", "Inferred from explicit type/sheet/source text into RISK/ISSUE/ACTION/ASSUMPTION/DEPENDENCY/OTHER.")
    add_metric_doc("PQ_RAID_Detail", "IsOpen", "Open flag", "True when status is not closed/resolved/complete/done/mitigated.")
    add_metric_doc("PQ_RAID_Detail", "IsClosed", "Closed flag", "Status contains closed/resolved/complete/done/mitigated/cancelled.")
    add_metric_doc("PQ_RAID_Detail", "IsOverdue", "Overdue flag", "IsOpen AND DueDate < today.")
    add_metric_doc("PQ_RAID_Detail", "IsDueNext14", "Due next 14 days flag", "IsOpen AND DueDate between today and today+14 days.")
    add_metric_doc("PQ_RAID_Detail", "IsHighPriority", "High-priority flag", "Priority text contains critical/high/p1/sev1 OR RiskScore >= configured threshold.")

    add_metric_doc("PQ_RAID_Summary", "TotalItems", "Total RAID items", "COUNT rows in PQ_RAID_Detail by RaidType (plus ALL rollup).")
    add_metric_doc("PQ_RAID_Summary", "OpenItems", "Open items", "SUM(IsOpen).")
    add_metric_doc("PQ_RAID_Summary", "ClosedItems", "Closed items", "SUM(IsClosed).")
    add_metric_doc("PQ_RAID_Summary", "OverdueItems", "Overdue items", "SUM(IsOverdue).")
    add_metric_doc("PQ_RAID_Summary", "DueNext14", "Due in next 14 days", "SUM(IsDueNext14).")
    add_metric_doc("PQ_RAID_Summary", "HighPriorityItems", "High-priority items", "SUM(IsHighPriority).")

    add_metric_doc("PQ_RAID_Workstream", "TotalItems", "Total items by workstream", "Aggregation by RaidType+Workstream (+Owner rows).")
    add_metric_doc("PQ_RAID_Workstream", "OpenItems", "Open items by workstream", "SUM(IsOpen) by grouping.")
    add_metric_doc("PQ_RAID_Workstream", "OverdueItems", "Overdue by workstream", "SUM(IsOverdue) by grouping.")
    add_metric_doc("PQ_RAID_Workstream", "HighPriorityItems", "High priority by workstream", "SUM(IsHighPriority) by grouping.")

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


def build_raid_domain(domain_config: dict[str, Any], source_frames: dict[str, list[SourceFrame]]) -> DomainBuildResult:
    source_names = domain_config.get("source_names", [])
    aliases: dict[str, list[str]] = domain_config.get("aliases", {})
    high_priority_score_threshold = float(domain_config.get("high_priority_score_threshold", 12))

    records: list[pd.DataFrame] = []
    domain_name = str(domain_config.get("name", "raid"))
    validation_rows = build_raid_validation_rows(domain_name)

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
                "PQ_RAID_Source": empty,
                "PQ_RAID_Detail": empty,
                "PQ_RAID_Summary": empty,
                "PQ_RAID_Workstream": empty,
            },
            validation_rows=validation_rows,
            validation_pq_order=[
                "PQ_RAID_Source",
                "PQ_RAID_Detail",
                "PQ_RAID_Summary",
                "PQ_RAID_Workstream",
            ],
        )

    source_df = pd.concat(records, ignore_index=True)

    record_id = _extract_series(source_df, aliases, "record_id").map(clean_text)
    task_id = _extract_series(source_df, aliases, "task_id").map(clean_text)
    defect_id = _extract_series(source_df, aliases, "defect_id").map(clean_text)
    title = _extract_series(source_df, aliases, "task_name").map(clean_text)
    workstream = _extract_series(source_df, aliases, "workstream").map(clean_text).replace("", "(Unassigned)")
    owner = _extract_series(source_df, aliases, "owner").map(clean_text).replace("", "(Unassigned)")
    raid_type_raw = _extract_series(source_df, aliases, "raid_type").map(clean_text)
    status = _extract_series(source_df, aliases, "status").map(clean_text)
    priority = _extract_series(source_df, aliases, "priority").map(clean_text)
    due_date = pd.to_datetime(_extract_series(source_df, aliases, "due_date"), errors="coerce")
    logged_by = _extract_series(source_df, aliases, "logged_by").map(clean_text)
    description = _extract_series(source_df, aliases, "description").map(clean_text)
    mitigation = _extract_series(source_df, aliases, "mitigation").map(clean_text)
    comments = _extract_series(source_df, aliases, "comments").map(clean_text)
    notes = _extract_series(source_df, aliases, "notes").map(clean_text)
    category = _extract_series(source_df, aliases, "category").map(clean_text)
    probability = _extract_series(source_df, aliases, "probability").map(clean_text)
    impact = _extract_series(source_df, aliases, "impact").map(clean_text)
    risk_score = as_numeric(_extract_series(source_df, aliases, "risk_score"))

    def _infer_status_series(df: pd.DataFrame) -> pd.Series:
        metadata_cols = {"source_name", "source_file", "source_sheet", "source_row_number"}
        pattern = r"\b(?:open|closed|in progress|in-progress|complete|completed|done|resolved|cancelled|canceled|mitigated|deferred|pending|active|new)\b"
        best_col: str | None = None
        best_score = -1
        best_ratio = 0.0
        for col in df.columns:
            col_name = clean_text(col)
            if col_name in metadata_cols:
                continue
            s = df[col].map(clean_text)
            non_blank = s.ne("")
            non_blank_count = int(non_blank.sum())
            if non_blank_count < 10:
                continue
            matches = s.str.lower().str.contains(pattern, regex=True)
            score = int(matches.sum())
            ratio = float(score / non_blank_count) if non_blank_count else 0.0
            if score > best_score or (score == best_score and ratio > best_ratio):
                best_col = str(col)
                best_score = score
                best_ratio = ratio
        if best_col is None or best_score <= 0 or best_ratio < 0.3:
            return pd.Series("", index=df.index)
        return df[best_col].map(clean_text)

    def _infer_due_date_series(df: pd.DataFrame) -> pd.Series:
        metadata_cols = {"source_name", "source_file", "source_sheet", "source_row_number"}
        parsed_cache: dict[str, pd.Series] = {}
        best_col: str | None = None
        best_score = -1.0

        for col in df.columns:
            col_name = clean_text(col)
            if col_name in metadata_cols:
                continue
            raw = df[col]
            parsed = pd.to_datetime(raw, errors="coerce", format="mixed")
            parsed_cache[str(col)] = parsed

            parseable = int(parsed.notna().sum())
            if parseable < 20:
                continue

            ratio = float(parseable / max(1, len(parsed)))
            years = parsed.dropna().dt.year
            if years.empty:
                continue
            plausible_year_ratio = float(((years >= 2020) & (years <= 2035)).mean())
            if plausible_year_ratio < 0.6:
                continue

            unique_vals = int(parsed.dropna().nunique())
            name_bonus = 0.0
            if re.search(r"due|target|date|deadline|end", col_name, re.IGNORECASE):
                name_bonus += 0.25

            # Prefer columns with meaningful variety and plausible date density.
            score = ratio + min(unique_vals / 500.0, 0.2) + name_bonus
            if score > best_score:
                best_col = str(col)
                best_score = score

        if best_col is None:
            return pd.Series(pd.NaT, index=df.index)
        return parsed_cache[best_col]

    # Fallback: if mapped status is mostly blank, infer status-like column from source values.
    mapped_non_blank = int(status.map(clean_text).ne("").sum())
    if mapped_non_blank < max(10, int(len(status) * 0.1)):
        inferred_status = _infer_status_series(source_df)
        if int(inferred_status.map(clean_text).ne("").sum()) > mapped_non_blank:
            status = inferred_status

    # Fallback: infer due/target date column when mapped due date is mostly empty.
    mapped_due_non_null = int(pd.Series(due_date).notna().sum())
    if mapped_due_non_null < max(10, int(len(source_df) * 0.1)):
        inferred_due_date = _infer_due_date_series(source_df)
        if int(pd.Series(inferred_due_date).notna().sum()) > mapped_due_non_null:
            due_date = inferred_due_date

    raid_text = (
        raid_type_raw.map(clean_text)
        + " | "
        + source_df["source_sheet"].map(clean_text)
        + " | "
        + source_df["source_name"].map(clean_text)
        + " | "
        + title.map(clean_text)
    ).str.lower()

    def infer_raid_type(text: str) -> str:
        t = clean_text(text).lower()
        if re.search(r"\brisk\b", t):
            return "RISK"
        if re.search(r"\bissue\b", t):
            return "ISSUE"
        if re.search(r"\baction\b", t):
            return "ACTION"
        if re.search(r"\bassumption\b", t):
            return "ASSUMPTION"
        if re.search(r"\bdependency\b", t):
            return "DEPENDENCY"
        return "OTHER"

    raid_type = raid_text.map(infer_raid_type)

    status_norm = status_text(status)
    closed_pattern = r"\b(?:closed|resolve|resolved|complete|completed|done|cancelled|canceled|mitigated|deferred|duplicate)\b"
    open_pattern = r"\b(?:open|active|new|in progress|in-progress|ongoing|pending)\b"
    is_closed = status_norm.str.contains(closed_pattern, regex=True)
    is_open = (~is_closed) & (status_norm.str.contains(open_pattern, regex=True) | status_norm.eq(""))

    today = pd.Timestamp.now().normalize()
    is_overdue = is_open & due_date.notna() & (due_date.dt.normalize() < today)
    is_due_next14 = is_open & due_date.notna() & (due_date.dt.normalize().between(today, today + pd.Timedelta(days=14)))

    priority_norm = priority.str.lower()
    is_high_priority = (
        priority_norm.str.contains(r"critical|high|p1|sev1", regex=True)
        | risk_score.fillna(0).ge(high_priority_score_threshold)
    )

    detail_df = pd.DataFrame(
        {
            "RaidType": raid_type,
            "RecordID": record_id.where(record_id != "", source_df["source_row_number"].map(lambda n: f"ROW-{n}")),
            "TaskID": task_id,
            "DefectID": defect_id,
            "Title": title,
            "Workstream": workstream,
            "Owner": owner,
            "Status": status,
            "Priority": priority,
            "DueDate": due_date,
            "LoggedBy": logged_by,
            "Description": description,
            "Mitigation": mitigation,
            "Comments": comments,
            "Notes": notes,
            "Category": category,
            "Probability": probability,
            "Impact": impact,
            "RiskScore": risk_score,
            "IsOpen": is_open,
            "IsClosed": is_closed,
            "IsOverdue": is_overdue,
            "IsDueNext14": is_due_next14,
            "IsHighPriority": is_high_priority,
            "SourceSheet": source_df["source_sheet"].map(clean_text),
            "SourceFile": source_df["source_file"].map(clean_text),
            "SourceRow": source_df["source_row_number"],
        }
    )

    # Guardrail: drop fully blank logical rows.
    meaningful = (
        detail_df["Title"].map(clean_text).ne("")
        | detail_df["RecordID"].map(clean_text).ne("")
        | detail_df["Status"].map(clean_text).ne("")
        | detail_df["Priority"].map(clean_text).ne("")
    )
    detail_df = detail_df.loc[meaningful].copy()

    def _aggregate(group_cols: list[str]) -> pd.DataFrame:
        agg_spec = {
            "TotalItems": ("RecordID", "count"),
            "OpenItems": ("IsOpen", "sum"),
            "ClosedItems": ("IsClosed", "sum"),
            "OverdueItems": ("IsOverdue", "sum"),
            "DueNext14": ("IsDueNext14", "sum"),
            "HighPriorityItems": ("IsHighPriority", "sum"),
        }
        if group_cols:
            return detail_df.groupby(group_cols, dropna=False).agg(**agg_spec).reset_index()
        return pd.DataFrame(
            [
                {
                    "TotalItems": int(detail_df["RecordID"].count()),
                    "OpenItems": int(detail_df["IsOpen"].sum()),
                    "ClosedItems": int(detail_df["IsClosed"].sum()),
                    "OverdueItems": int(detail_df["IsOverdue"].sum()),
                    "DueNext14": int(detail_df["IsDueNext14"].sum()),
                    "HighPriorityItems": int(detail_df["IsHighPriority"].sum()),
                }
            ]
        )

    summary_by_type = _aggregate(["RaidType"])
    summary_all = _aggregate([])
    if not summary_all.empty:
        summary_all.insert(0, "RaidType", "ALL")
    summary_df = pd.concat([summary_all, summary_by_type], ignore_index=True)

    ws_rollup = _aggregate(["RaidType", "Workstream"])
    ws_rollup.insert(2, "Owner", "(All Owners)")
    ws_owner = _aggregate(["RaidType", "Workstream", "Owner"])
    workstream_df = pd.concat([ws_rollup, ws_owner], ignore_index=True)

    detail_df = detail_df.sort_values(["RaidType", "IsOverdue", "IsHighPriority", "DueDate", "RecordID"], ascending=[True, False, False, True, True], kind="stable")
    summary_df = summary_df.sort_values(["RaidType"], kind="stable")
    workstream_df = workstream_df.sort_values(["RaidType", "Workstream", "Owner"], kind="stable")

    return DomainBuildResult(
        tables={
            "PQ_RAID_Source": source_df,
            "PQ_RAID_Detail": detail_df,
            "PQ_RAID_Summary": summary_df,
            "PQ_RAID_Workstream": workstream_df,
        },
        validation_rows=validation_rows,
        validation_pq_order=[
            "PQ_RAID_Source",
            "PQ_RAID_Detail",
            "PQ_RAID_Summary",
            "PQ_RAID_Workstream",
        ],
    )


def build_tasks_domain(domain_config: dict[str, Any], source_frames: dict[str, list[SourceFrame]]) -> DomainBuildResult:
    """Build Tasks domain from MS Project file.
    
    Reads .mpp file and applies hierarchical filtering:
    - Filters out leaf tasks (Outline Level 6+)
    - Computes immediate-attention flags
    - Creates canonical PQ_Task_* tables
    """
    import win32com.client
    
    mpp_path = domain_config.get("mpp_path", "")
    if not mpp_path or not Path(mpp_path).exists():
        print(f"[warning] tasks domain: mpp_path not found: {mpp_path}")
        empty = pd.DataFrame()
        return DomainBuildResult(
            tables={
                "PQ_Task_Source": empty,
                "PQ_Task_Detail": empty,
                "PQ_Task_Summary": empty,
                "PQ_Task_Workstream": empty,
            },
            validation_rows=[],
            validation_pq_order=[
                "PQ_Task_Source",
                "PQ_Task_Detail",
                "PQ_Task_Summary",
                "PQ_Task_Workstream",
            ],
        )
    
    try:
        # Open MS Project file without UI
        msp = win32com.client.Dispatch("MSProject.Application")
        msp.Visible = False
        msp.DisplayAlerts = False
        msp.FileOpen(str(mpp_path), True, 0)  # read-only, no save
        
        proj = msp.ActiveProject
        tasks_data = []
        
        # Extract all tasks, including candidate custom text fields used for Workstream.
        for task in proj.Tasks:
            if task is None:
                continue
            
            try:
                parent_id = task.OutlineParent.ID if task.OutlineLevel > 1 else None
            except:
                parent_id = None
            
            try:
                start_date = str(task.Start)[:10] if task.Start else None
                finish_date = str(task.Finish)[:10] if task.Finish else None
            except:
                start_date = None
                finish_date = None
            
            try:
                pct_complete = int(task.PercentComplete) if task.PercentComplete else 0
            except:
                pct_complete = 0

            try:
                is_summary_task = bool(task.Summary)
            except:
                is_summary_task = False

            try:
                is_milestone_task = bool(task.Milestone)
            except:
                is_milestone_task = False

            try:
                is_active_task = bool(task.Active)
            except:
                is_active_task = True

            text_fields: dict[str, str] = {}
            for i in range(1, 11):
                attr = f"Text{i}"
                try:
                    text_fields[attr] = clean_text(getattr(task, attr))
                except Exception:
                    text_fields[attr] = ""
            
            tasks_data.append({
                'ID': task.ID,
                'Name': task.Name if task.Name else "",
                'OutlineLevel': task.OutlineLevel,
                'ParentID': parent_id,
                'Start': start_date,
                'Finish': finish_date,
                'PercentComplete': pct_complete,
                'UniqueID': task.UniqueID,
                'IsSummaryTask': is_summary_task,
                'IsMilestoneTask': is_milestone_task,
                'IsActiveTask': is_active_task,
                **text_fields,
            })
        
        msp.FileCloseAll(2)  # pjDoNotSave
        msp.Quit()
        
        if not tasks_data:
            empty = pd.DataFrame()
            return DomainBuildResult(
                tables={
                    "PQ_Task_Source": empty,
                    "PQ_Task_Detail": empty,
                    "PQ_Task_Summary": empty,
                    "PQ_Task_Workstream": empty,
                    "PQ_Task_Milestone": empty,
                },
                validation_rows=[],
                validation_pq_order=[
                    "PQ_Task_Source",
                    "PQ_Task_Detail",
                    "PQ_Task_Summary",
                    "PQ_Task_Workstream",
                    "PQ_Task_Milestone",
                ],
            )
        
        source_df = pd.DataFrame(tasks_data)
        source_df.insert(0, "source_file", str(mpp_path))
        
        # Workstream must come from Workstream field (MS Project custom text field).
        # Configurable; defaults to Text1. Falls back to densest configured candidate.
        configured_workstream_field = clean_text(domain_config.get("workstream_field", "Text1")) or "Text1"
        workstream_field_candidates = domain_config.get(
            "workstream_field_candidates",
            ["Text1", "Text2", "Text3", "Text4", "Text5", "Text6", "Text7", "Text8", "Text9", "Text10"],
        )
        workstream_field_candidates = [clean_text(c) for c in workstream_field_candidates if clean_text(c) != ""]
        if configured_workstream_field not in workstream_field_candidates:
            workstream_field_candidates.insert(0, configured_workstream_field)

        chosen_workstream_field = configured_workstream_field if configured_workstream_field in source_df.columns else ""
        if chosen_workstream_field == "" or int(source_df.get(chosen_workstream_field, pd.Series(dtype=object)).map(clean_text).ne("").sum()) == 0:
            best_col = ""
            best_non_blank = -1
            for col in workstream_field_candidates:
                if col not in source_df.columns:
                    continue
                non_blank = int(source_df[col].map(clean_text).ne("").sum())
                if non_blank > best_non_blank:
                    best_col = col
                    best_non_blank = non_blank
            chosen_workstream_field = best_col

        if chosen_workstream_field != "":
            workstream_raw = source_df[chosen_workstream_field].map(clean_text)
        else:
            workstream_raw = pd.Series("", index=source_df.index)

        # Count execution tasks only:
        # - exclude Close4D milestone names
        # - exclude summary tasks
        # - exclude milestone tasks
        # - exclude blank task names
        name_norm = source_df['Name'].map(clean_text)
        name_norm_lower = name_norm.str.lower()

        is_close4d = name_norm_lower.str.contains(r"close\s*4d|close4d", regex=True)
        is_summary = source_df['IsSummaryTask'].fillna(False).astype(bool) | name_norm_lower.str.contains(r"\bsummary\b", regex=True)
        is_milestone = source_df['IsMilestoneTask'].fillna(False).astype(bool)
        is_blank_name = name_norm.eq("")

        is_execution_task = (~is_close4d) & (~is_summary) & (~is_milestone) & (~is_blank_name)
        filtered_df = source_df[is_execution_task].copy()
        
        if filtered_df.empty:
            empty = pd.DataFrame()
            return DomainBuildResult(
                tables={
                    "PQ_Task_Source": source_df,
                    "PQ_Task_Detail": empty,
                    "PQ_Task_Summary": empty,
                    "PQ_Task_Workstream": empty,
                    "PQ_Task_Milestone": empty,
                },
                validation_rows=[],
                validation_pq_order=[
                    "PQ_Task_Source",
                    "PQ_Task_Detail",
                    "PQ_Task_Summary",
                    "PQ_Task_Workstream",
                    "PQ_Task_Milestone",
                ],
            )
        
        # Extract task attributes
        task_id = filtered_df['ID'].astype(str)
        task_name = filtered_df['Name'].map(clean_text)
        outline_level = filtered_df['OutlineLevel'].astype(int)
        parent_id = filtered_df['ParentID'].fillna(0).astype(int)
        start_date = pd.to_datetime(filtered_df['Start'], errors='coerce')
        finish_date = pd.to_datetime(filtered_df['Finish'], errors='coerce')
        pct_complete = filtered_df['PercentComplete'].astype(float)

        # Workstream comes from explicit Workstream field only.
        workstream = workstream_raw.reindex(filtered_df.index).map(clean_text).replace("", "(Unassigned)")

        # Derive milestone from ancestor hierarchy (prefer level-3 ancestor, fallback level-2).
        source_lookup = source_df.set_index("ID", drop=False)

        def derive_milestone(row: pd.Series) -> str:
            parent = int(row.get("ParentID") or 0)
            fallback_l2 = ""
            safety = 0
            while parent > 0 and safety < 30:
                safety += 1
                if parent not in source_lookup.index:
                    break
                p = source_lookup.loc[parent]
                pname = clean_text(p.get("Name"))
                try:
                    plvl = int(p.get("OutlineLevel"))
                except Exception:
                    plvl = -1
                if plvl == 3 and pname != "":
                    return pname
                if plvl == 2 and pname != "" and fallback_l2 == "":
                    fallback_l2 = pname
                try:
                    parent = int(p.get("ParentID") or 0)
                except Exception:
                    break
            return fallback_l2 or "(Unassigned)"

        milestone = filtered_df.apply(derive_milestone, axis=1)
        
        # Compute task status
        today = pd.Timestamp.now().normalize()
        
        # Status logic:
        # - If % Complete = 100 → Complete
        # - If % Complete > 0 and < 100 → In Progress
        # - If Finish date < today and % Complete < 100 → Overdue
        # - Otherwise → Not Started
        def compute_status(idx: int) -> str:
            pct = pct_complete.iloc[idx]
            finish = finish_date.iloc[idx]
            
            if pct >= 100:
                return "Complete"
            if pct > 0:
                return "In Progress"
            if pd.notna(finish) and finish.normalize() < today:
                return "Overdue"
            return "Not Started"
        
        status = pd.Series([compute_status(i) for i in range(len(filtered_df))], index=filtered_df.index)
        
        # Flags
        is_complete = pct_complete >= 100
        is_in_progress = (pct_complete > 0) & (pct_complete < 100)
        is_open = ~is_complete
        is_overdue = is_open & finish_date.notna() & (finish_date.dt.normalize() < today)
        is_due_next14 = is_open & finish_date.notna() & (finish_date.dt.normalize().between(today, today + pd.Timedelta(days=14)))
        
        # Immediate Attention: Open tasks that are:
        # - Overdue, OR
        # - Due in next 7 days AND > 0% progress (showing activity), OR
        # - In progress but past target date
        is_due_next7 = is_open & finish_date.notna() & (finish_date.dt.normalize().between(today, today + pd.Timedelta(days=7)))
        is_immediate_attention = is_overdue | (is_due_next7 & is_in_progress) | (is_in_progress & is_overdue)
        is_potential_risk = is_immediate_attention | (is_due_next14 & is_open)
        
        reason_items = []
        for idx_counter, _ in enumerate(filtered_df.itertuples(index=False)):
            reasons = []
            if is_overdue.iloc[idx_counter]:
                reasons.append("Overdue")
            if is_due_next7.iloc[idx_counter] and is_in_progress.iloc[idx_counter]:
                reasons.append("Due Soon (In Progress)")
            if is_in_progress.iloc[idx_counter]:
                if pct_complete.iloc[idx_counter] < 50:
                    reasons.append("Low Progress")
            reason_items.append(" | ".join(reasons) if reasons else "")
        
        reason = pd.Series(reason_items, index=filtered_df.index)
        
        detail_df = pd.DataFrame({
            'TaskID': task_id,
            'TaskName': task_name,
            'OutlineLevel': outline_level,
            'ParentID': parent_id,
            'Workstream': workstream,
            'Milestone': milestone,
            'WorkstreamField': chosen_workstream_field,
            'Status': status,
            'PercentComplete': pct_complete,
            'StartDate': start_date,
            'FinishDate': finish_date,
            'IsOpen': is_open,
            'IsComplete': is_complete,
            'IsInProgress': is_in_progress,
            'IsOverdue': is_overdue,
            'IsDueNext14': is_due_next14,
            'IsDueNext7': is_due_next7,
            'IsImmediateAttention': is_immediate_attention,
            'IsPotentialRisk': is_potential_risk,
            'IsExecutionTask': True,
            'IsSummaryTask': filtered_df['IsSummaryTask'].fillna(False).astype(bool).values,
            'IsMilestoneTask': filtered_df['IsMilestoneTask'].fillna(False).astype(bool).values,
            'IsClose4D': name_norm_lower.reindex(filtered_df.index).str.contains(r"close\s*4d|close4d", regex=True).fillna(False).values,
            'Reason': reason,
            'SourceFile': filtered_df['source_file'].map(clean_text),
        })
        
        # Clean: drop rows with no task name
        detail_df = detail_df[detail_df['TaskName'].ne("")]
        
        # Summary aggregation
        def _aggregate_tasks(group_cols: list[str]) -> pd.DataFrame:
            agg_spec = {
                'TotalTasks': ('TaskID', 'count'),
                'OpenTasks': ('IsOpen', 'sum'),
                'CompleteTasks': ('IsComplete', 'sum'),
                'InProgressTasks': ('IsInProgress', 'sum'),
                'OverdueTasks': ('IsOverdue', 'sum'),
                'DueNext14': ('IsDueNext14', 'sum'),
                'DueNext7': ('IsDueNext7', 'sum'),
                'ImmediateAttention': ('IsImmediateAttention', 'sum'),
                'PotentialRiskTasks': ('IsPotentialRisk', 'sum'),
            }
            if group_cols:
                return detail_df.groupby(group_cols, dropna=False).agg(**agg_spec).reset_index()
            return pd.DataFrame([{
                'TotalTasks': int(detail_df['TaskID'].count()),
                'OpenTasks': int(detail_df['IsOpen'].sum()),
                'CompleteTasks': int(detail_df['IsComplete'].sum()),
                'InProgressTasks': int(detail_df['IsInProgress'].sum()),
                'OverdueTasks': int(detail_df['IsOverdue'].sum()),
                'DueNext14': int(detail_df['IsDueNext14'].sum()),
                'DueNext7': int(detail_df['IsDueNext7'].sum()),
                'ImmediateAttention': int(detail_df['IsImmediateAttention'].sum()),
                'PotentialRiskTasks': int(detail_df['IsPotentialRisk'].sum()),
            }])
        
        summary_by_workstream = _aggregate_tasks(['Workstream'])
        summary_all = _aggregate_tasks([])
        if not summary_all.empty:
            summary_all.insert(0, 'Workstream', 'ALL')
        summary_df = pd.concat([summary_all, summary_by_workstream], ignore_index=True)
        
        # Workstream detail (by status)
        ws_by_status = _aggregate_tasks(['Workstream', 'Status'])
        workstream_df = ws_by_status.copy()

        # Milestone rollup (ALL milestones only).
        milestone_df = _aggregate_tasks(['Milestone'])
        
        detail_df = detail_df.sort_values(['IsImmediateAttention', 'IsOverdue', 'IsDueNext7', 'FinishDate', 'TaskID'], 
                                          ascending=[False, False, False, True, True], kind='stable')
        summary_df = summary_df.sort_values(['Workstream'], kind='stable')
        workstream_df = workstream_df.sort_values(['Workstream', 'Status'], kind='stable')
        milestone_df = milestone_df.sort_values(['OpenTasks', 'Milestone'], ascending=[False, True], kind='stable')
        
        return DomainBuildResult(
            tables={
                'PQ_Task_Source': source_df,
                'PQ_Task_Detail': detail_df,
                'PQ_Task_Summary': summary_df,
                'PQ_Task_Workstream': workstream_df,
                'PQ_Task_Milestone': milestone_df,
            },
            validation_rows=[],
            validation_pq_order=[
                'PQ_Task_Source',
                'PQ_Task_Detail',
                'PQ_Task_Summary',
                'PQ_Task_Workstream',
                'PQ_Task_Milestone',
            ],
        )
    
    except Exception as e:
        print(f"[error] tasks domain: {e}")
        import traceback
        traceback.print_exc()
        try:
            msp.FileCloseAll(2)
            msp.Quit()
        except:
            pass
        empty = pd.DataFrame()
        return DomainBuildResult(
            tables={
                'PQ_Task_Source': empty,
                'PQ_Task_Detail': empty,
                'PQ_Task_Summary': empty,
                'PQ_Task_Workstream': empty,
                'PQ_Task_Milestone': empty,
            },
            validation_rows=[],
            validation_pq_order=[
                'PQ_Task_Source',
                'PQ_Task_Detail',
                'PQ_Task_Summary',
                'PQ_Task_Workstream',
                'PQ_Task_Milestone',
            ],
        )


DOMAIN_BUILDERS: dict[str, Callable[[dict[str, Any], dict[str, list[SourceFrame]]], DomainBuildResult]] = {
    "testing_defects": build_testing_defects_domain,
    "raid": build_raid_domain,
    "tasks": build_tasks_domain,
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


def build_action_items_tables(all_tables: dict[str, pd.DataFrame]) -> dict[str, pd.DataFrame]:
    """Build combined action-items reporting tables from RAID, Tasks, and Testing/Defects domains."""

    def _bool_series(df: pd.DataFrame, col: str) -> pd.Series:
        if col not in df.columns:
            return pd.Series(False, index=df.index)
        s = df[col]
        if str(getattr(s, "dtype", "")) == "bool":
            return s.fillna(False)
        return s.map(lambda v: str(v).strip().lower() in {"true", "1", "yes", "y"})

    def _text_series(df: pd.DataFrame, col: str, default: str = "") -> pd.Series:
        if col not in df.columns:
            return pd.Series(default, index=df.index)
        return df[col].map(clean_text)

    today = pd.Timestamp.now().normalize()
    in_two_weeks = today + pd.Timedelta(days=14)

    combined_parts: list[pd.DataFrame] = []

    # RAID
    raid = all_tables.get("PQ_RAID_Detail", pd.DataFrame()).copy()
    if not raid.empty:
        workstream = _text_series(raid, "Workstream", "(Unassigned)").replace("", "(Unassigned)")
        assigned = _text_series(raid, "Owner", "(Unassigned)").replace("", "(Unassigned)")
        item_id = _text_series(raid, "RecordID")
        item_name = _text_series(raid, "Title")
        status = _text_series(raid, "Status")
        due_date = pd.to_datetime(raid.get("DueDate"), errors="coerce")
        raid_type = _text_series(raid, "RaidType", "OTHER").replace("", "OTHER")

        is_open = _bool_series(raid, "IsOpen")
        is_overdue = _bool_series(raid, "IsOverdue")
        is_high = _bool_series(raid, "IsHighPriority")

        status_lower = status.str.lower()
        needs_clar = (
            assigned.str.lower().isin(["", "(unassigned)", "unassigned", "unknown", "tbd"])
            | _text_series(raid, "Description").eq("")
            | _text_series(raid, "Mitigation").eq("")
            | _text_series(raid, "Comments").eq("")
            | _text_series(raid, "Notes").eq("")
            | status_lower.str.contains(r"clarif|need info|needs info|tbd|unknown", regex=True)
        )
        due_next_14 = is_open & due_date.notna() & due_date.dt.normalize().between(today, in_two_weeks)
        in_progress_status = status_lower.str.contains(r"in\s*progress|pending|active|ongoing", regex=True)

        is_immediate = is_open & (is_overdue | is_high)
        is_needs_soon = is_open & (~is_immediate) & (due_next_14 | needs_clar)
        is_in_progress_2w = is_open & (~is_immediate) & (in_progress_status | due_next_14)

        part = pd.DataFrame(
            {
                "SourceDomain": "RAID",
                "ActionType": "RAID-" + raid_type,
                "Workstream": workstream,
                "AssignedTo": assigned,
                "ItemID": item_id,
                "ItemName": item_name,
                "Status": status,
                "DueDate": due_date,
                "IsOpen": is_open,
                "IsInProgress": in_progress_status,
                "IsImmediate": is_immediate,
                "IsNeedsAttentionSoon": is_needs_soon,
                "IsInProgress2Weeks": is_in_progress_2w,
            }
        )
        combined_parts.append(part)

    # TASKS
    tasks = all_tables.get("PQ_Task_Detail", pd.DataFrame()).copy()
    if not tasks.empty:
        workstream = _text_series(tasks, "Workstream", "(Unassigned)").replace("", "(Unassigned)")
        assigned = pd.Series("(Unassigned)", index=tasks.index)
        item_id = _text_series(tasks, "TaskID")
        item_name = _text_series(tasks, "TaskName")
        status = _text_series(tasks, "Status")
        due_date = pd.to_datetime(tasks.get("FinishDate"), errors="coerce")

        is_open = _bool_series(tasks, "IsOpen")
        is_in_progress = _bool_series(tasks, "IsInProgress")
        is_immediate = _bool_series(tasks, "IsImmediateAttention")
        due_next_14 = _bool_series(tasks, "IsDueNext14")

        is_needs_soon = is_open & (~is_immediate) & due_next_14
        is_in_progress_2w = is_open & (~is_immediate) & (is_in_progress | due_next_14)

        part = pd.DataFrame(
            {
                "SourceDomain": "TASKS",
                "ActionType": "TASK",
                "Workstream": workstream,
                "AssignedTo": assigned,
                "ItemID": item_id,
                "ItemName": item_name,
                "Status": status,
                "DueDate": due_date,
                "IsOpen": is_open,
                "IsInProgress": is_in_progress,
                "IsImmediate": is_immediate,
                "IsNeedsAttentionSoon": is_needs_soon,
                "IsInProgress2Weeks": is_in_progress_2w,
            }
        )
        combined_parts.append(part)

    # TESTING / DEFECTS
    td = all_tables.get("PQ_TD_Detail", pd.DataFrame()).copy()
    if not td.empty:
        workstream = _text_series(td, "Workstream", "(Unassigned)").replace("", "(Unassigned)")
        assigned = _text_series(td, "Owner", "(Unassigned)").replace("", "(Unassigned)")
        item_id = _text_series(td, "TestID").where(_text_series(td, "TestID").ne(""), _text_series(td, "DefectID"))
        item_name = _text_series(td, "TestName").where(_text_series(td, "TestName").ne(""), _text_series(td, "TaskName"))
        status = _text_series(td, "Status")
        due_date = pd.to_datetime(td.get("PlannedDate"), errors="coerce")

        status_lower = status.str.lower()
        is_passed = _bool_series(td, "IsPassed")
        is_failed = _bool_series(td, "IsFailed")
        is_blocked = _bool_series(td, "IsBlocked")
        is_in_progress = _bool_series(td, "IsInProgress")
        is_not_run = _bool_series(td, "IsNotRun")

        is_open = ~is_passed
        due_next_14 = due_date.notna() & due_date.dt.normalize().between(today, in_two_weeks)
        has_defect_link = _bool_series(td, "HasDefectLink")

        is_immediate = is_open & (is_failed | is_blocked)
        is_needs_soon = is_open & (~is_immediate) & (due_next_14 | has_defect_link)
        is_in_progress_2w = is_open & (~is_immediate) & (is_in_progress | (is_not_run & due_next_14) | status_lower.str.contains(r"pending|queued", regex=True))

        part = pd.DataFrame(
            {
                "SourceDomain": "TD",
                "ActionType": "TD",
                "Workstream": workstream,
                "AssignedTo": assigned,
                "ItemID": item_id,
                "ItemName": item_name,
                "Status": status,
                "DueDate": due_date,
                "IsOpen": is_open,
                "IsInProgress": is_in_progress,
                "IsImmediate": is_immediate,
                "IsNeedsAttentionSoon": is_needs_soon,
                "IsInProgress2Weeks": is_in_progress_2w,
            }
        )
        combined_parts.append(part)

    if not combined_parts:
        empty = pd.DataFrame()
        return {
            "PQ_Action_Items_Detail": empty,
            "PQ_Action_Items_Summary": empty,
        }

    detail = pd.concat(combined_parts, ignore_index=True)
    detail["Workstream"] = detail["Workstream"].map(clean_text).replace("", "(Unassigned)")
    detail["AssignedTo"] = detail["AssignedTo"].map(clean_text).replace("", "(Unassigned)")
    detail["ItemID"] = detail["ItemID"].map(clean_text)
    detail["ItemName"] = detail["ItemName"].map(clean_text)
    detail["Status"] = detail["Status"].map(clean_text)

    detail = detail[
        detail["IsImmediate"].fillna(False)
        | detail["IsNeedsAttentionSoon"].fillna(False)
        | detail["IsInProgress2Weeks"].fillna(False)
    ].copy()

    def _section_label(row: pd.Series) -> str:
        if bool(row.get("IsImmediate", False)):
            return "Immediate"
        if bool(row.get("IsNeedsAttentionSoon", False)):
            return "Needs Attention Soon"
        if bool(row.get("IsInProgress2Weeks", False)):
            return "In Progress +2wks"
        return "Other"

    detail["Section"] = detail.apply(_section_label, axis=1)
    section_sort = {
        "Immediate": 1,
        "Needs Attention Soon": 2,
        "In Progress +2wks": 3,
        "Other": 9,
    }
    detail["SectionSort"] = detail["Section"].map(lambda v: section_sort.get(str(v), 9))

    detail = detail.sort_values(
        ["SectionSort", "Workstream", "AssignedTo", "ActionType", "DueDate", "ItemID"],
        ascending=[True, True, True, True, True, True],
        kind="stable",
    ).reset_index(drop=True)

    summary = (
        detail.groupby(["Workstream", "AssignedTo", "ActionType"], dropna=False)
        .agg(
            TotalActionItems=("ItemID", "count"),
            ImmediateItems=("IsImmediate", "sum"),
            NeedsAttentionSoonItems=("IsNeedsAttentionSoon", "sum"),
            InProgress2WeeksItems=("IsInProgress2Weeks", "sum"),
        )
        .reset_index()
        .sort_values(["Workstream", "AssignedTo", "TotalActionItems"], ascending=[True, True, False], kind="stable")
    )

    return {
        "PQ_Action_Items_Detail": detail,
        "PQ_Action_Items_Summary": summary,
    }


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

    # Cross-domain unified action items reporting tables.
    action_item_tables = build_action_items_tables(all_tables)
    all_tables.update(action_item_tables)

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

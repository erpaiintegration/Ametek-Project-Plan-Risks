/* eslint-disable no-console */
require("dotenv/config");

const fs = require("node:fs");
const path = require("node:path");
const crypto = require("node:crypto");
const { parse } = require("csv-parse/sync");
const { Client } = require("@notionhq/client");

const ROOT = process.cwd();
const STAGING_DIR = path.join(ROOT, "imports", "staging");
const EXPORTS_DIR = path.join(ROOT, "imports", "exports");

const NOTES_OWNER = (process.env.NOTES_OWNER || "notion").toLowerCase();
const PROJECTS_DB_ID = process.env.NOTION_PROJECTS_DB_ID;
const PROJECT_PAGE_ID = process.env.NOTION_AMETEK_PROJECT_PAGE_ID;
const PROJECT_IMPORT_VALUE = process.env.NOTION_PROJECT_IMPORT_VALUE;

const SUPPLEMENTAL_PLAN_NAME = "Project plan with resources and busines.csv";
const TASK_WORKSTREAM_FIELD = "Workstream";

const notion = process.env.NOTION_API_KEY
  ? new Client({ auth: process.env.NOTION_API_KEY })
  : null;

function ensureFolders() {
  fs.mkdirSync(STAGING_DIR, { recursive: true });
  fs.mkdirSync(EXPORTS_DIR, { recursive: true });
}

function getArgValue(flag) {
  const idx = process.argv.indexOf(flag);
  if (idx === -1) return undefined;
  return process.argv[idx + 1];
}

function listStagingCsvFiles() {
  return fs
    .readdirSync(STAGING_DIR)
    .filter((name) => name.toLowerCase().endsWith(".csv"))
    .map((name) => ({
      name,
      full: path.join(STAGING_DIR, name),
      mtime: fs.statSync(path.join(STAGING_DIR, name)).mtimeMs
    }))
    .sort((a, b) => b.mtime - a.mtime);
}

function parseCsvRows(filePath) {
  const raw = fs.readFileSync(filePath, "utf8");
  const rows = parse(raw, {
    columns: true,
    bom: true,
    skip_empty_lines: true,
    trim: true,
    relax_column_count: true,
    group_columns_by_name: true
  });
  return { raw, rows };
}

function detectTaskCsvScore(filePath) {
  try {
    const { rows } = parseCsvRows(filePath);
    if (!rows.length) return 0;
    const headers = Object.keys(rows[0]);
    const hasUniqueId = headers.includes("Unique ID");
    const hasTaskName = headers.includes("Task Name") || headers.includes("Task description");
    const hasWbs = headers.includes("WBS") || headers.includes("Outline Number");
    if (!(hasUniqueId && hasTaskName && hasWbs)) return 0;

    let score = 10;
    if (filePath.toLowerCase().includes("msp task")) score += 10;
    if (filePath.toLowerCase().includes("replan")) score += 5;
    return score;
  } catch {
    return 0;
  }
}

function pickPrimaryTaskCsv() {
  const explicit = getArgValue("--file");
  if (explicit) {
    const full = path.isAbsolute(explicit) ? explicit : path.join(ROOT, explicit);
    if (!fs.existsSync(full)) {
      throw new Error(`CSV file not found: ${full}`);
    }
    return full;
  }

  const files = listStagingCsvFiles();
  if (!files.length) {
    throw new Error(`No CSV files found in ${STAGING_DIR}.`);
  }

  const scored = files
    .map((f) => ({ ...f, score: detectTaskCsvScore(f.full) }))
    .filter((f) => f.score > 0)
    .sort((a, b) => b.score - a.score || b.mtime - a.mtime);

  if (!scored.length) {
    throw new Error(
      "Could not find a valid task CSV in staging. Ensure a file has Unique ID, Task Name, and WBS/Outline Number."
    );
  }

  return scored[0].full;
}

function findOptionalStagingFile(targetName) {
  const files = listStagingCsvFiles();
  const hit = files.find((f) => f.name.toLowerCase() === targetName.toLowerCase());
  return hit?.full;
}

function parseSupplementalPlan(filePath) {
  if (!filePath || !fs.existsSync(filePath)) {
    return new Map();
  }

  const raw = fs.readFileSync(filePath, "utf8");
  const matrix = parse(raw, {
    columns: false,
    bom: true,
    skip_empty_lines: false,
    trim: true,
    relax_column_count: true
  });

  const headerIndex = matrix.findIndex(
    (row) => Array.isArray(row) && row.includes("Unique ID") && row.includes("Task description")
  );
  if (headerIndex === -1) {
    return new Map();
  }

  const header = matrix[headerIndex];
  const idx = {
    uniqueId: header.indexOf("Unique ID"),
    resources: header.indexOf("Resources"),
    businessValidation: header.indexOf("Business Validation"),
    workstream: header.indexOf("Workstream"),
    taskDescription: header.indexOf("Task description")
  };

  const out = new Map();
  for (let i = headerIndex + 2; i < matrix.length; i += 1) {
    const row = matrix[i] || [];
    const uid = parseNumber(row[idx.uniqueId]);
    if (uid == null) continue;

    out.set(uid, {
      taskName: String(row[idx.taskDescription] || "").trim(),
      resourceNames: String(row[idx.resources] || "").trim(),
      businessValidationOwner: String(row[idx.businessValidation] || "").trim(),
      workstream: String(row[idx.workstream] || "").trim()
    });
  }

  return out;
}

function firstValue(row, key) {
  const raw = row[key];
  if (Array.isArray(raw)) {
    for (const value of raw) {
      if (value == null) continue;
      const text = String(value).trim();
      if (text.length > 0) return text;
    }
    return "";
  }
  if (raw == null) return "";
  return String(raw).trim();
}

function parseNumber(value) {
  if (!value) return null;
  const cleaned = String(value)
    .replace(/,/g, "")
    .replace(/%/g, "")
    .trim();
  if (cleaned.length === 0) return null;
  const num = Number(cleaned);
  return Number.isFinite(num) ? num : null;
}

function parseDate(value) {
  if (!value) return "";
  const d = new Date(String(value).trim());
  if (Number.isNaN(d.getTime())) return "";
  const mm = String(d.getMonth() + 1).padStart(2, "0");
  const dd = String(d.getDate()).padStart(2, "0");
  const yyyy = String(d.getFullYear());
  return `${mm}-${dd}-${yyyy}`;
}

function parseBoolYesNo(value) {
  if (!value) return false;
  const normalized = String(value).trim().toLowerCase();
  return ["yes", "true", "1", "y"].includes(normalized);
}

function parseIdList(value) {
  if (!value) return [];
  const matches = String(value).match(/\d+/g);
  if (!matches) return [];
  const ids = [];
  for (const token of matches) {
    const num = Number(token);
    if (Number.isFinite(num)) ids.push(num);
  }
  return [...new Set(ids)];
}

function deriveStatusUpdate(percentComplete) {
  const pct = Number(percentComplete);
  if (!Number.isFinite(pct)) return "Not started";
  if (pct >= 100) return "Done";
  if (pct >= 25) return "In progress";
  return "Not started";
}

function parentHierarchyKey(key) {
  if (!key) return "";
  const parts = key.split(".").map((x) => x.trim()).filter(Boolean);
  if (parts.length <= 1) return "";
  return parts.slice(0, -1).join(".");
}

function sha1Text(value) {
  return crypto.createHash("sha1").update(value).digest("hex");
}

function buildRowHash(row) {
  const payload = {
    "Task Name": firstValue(row, "Task Name"),
    WBS: firstValue(row, "WBS"),
    "Outline Number": firstValue(row, "Outline Number"),
    "Outline Level": firstValue(row, "Outline Level"),
    Summary: firstValue(row, "Summary"),
    "Task Mode": firstValue(row, "Task Mode"),
    "% Complete": firstValue(row, "% Complete"),
    Duration: firstValue(row, "Duration"),
    Start: firstValue(row, "Start"),
    Finish: firstValue(row, "Finish"),
    "Actual Start": firstValue(row, "Actual Start"),
    "Baseline10 Start": firstValue(row, "Baseline10 Start"),
    "Baseline10 Finish": firstValue(row, "Baseline10 Finish"),
    "Baseline Start": firstValue(row, "Baseline Start"),
    "Baseline Finish": firstValue(row, "Baseline Finish"),
    "Constraint Type": firstValue(row, "Constraint Type"),
    "Constraint Date": firstValue(row, "Constraint Date"),
    Critical: firstValue(row, "Critical"),
    Milestone: firstValue(row, "Milestone"),
    "Total Slack": firstValue(row, "Total Slack"),
    "Free Slack": firstValue(row, "Free Slack"),
    Predecessors: firstValue(row, "Predecessors"),
    "Unique ID Predecessors": firstValue(row, "Unique ID Predecessors"),
    "Unique ID Successors": firstValue(row, "Unique ID Successors"),
    "WBS Predecessors": firstValue(row, "WBS Predecessors"),
    "WBS Successors": firstValue(row, "WBS Successors"),
    "Resource Names": firstValue(row, "Resource Names"),
    "Business Validation Owner": firstValue(row, "Business Validation Owner"),
    Workstream: firstValue(row, TASK_WORKSTREAM_FIELD),
    Status: firstValue(row, "Status"),
    Warning: firstValue(row, "Warning"),
    "Error Message": firstValue(row, "Error Message"),
    Active: firstValue(row, "Active"),
    Placeholder: firstValue(row, "Placeholder"),
    "Linked Fields": firstValue(row, "Linked Fields")
  };

  if (NOTES_OWNER === "msp") {
    payload.Notes = firstValue(row, "Notes");
  }

  return sha1Text(JSON.stringify(payload));
}

function boolToText(v) {
  return v ? "TRUE" : "FALSE";
}

function csvEscape(value) {
  const text = String(value ?? "");
  if (/[",\n\r]/.test(text)) {
    return `"${text.replace(/"/g, '""')}"`;
  }
  return text;
}

function writeCsv(filePath, headers, rows) {
  const lines = [];
  lines.push(headers.map(csvEscape).join(","));
  for (const row of rows) {
    lines.push(headers.map((h) => csvEscape(row[h] ?? "")).join(","));
  }
  fs.writeFileSync(filePath, lines.join("\n"), "utf8");
}

async function resolveProjectImportValue() {
  if (PROJECT_IMPORT_VALUE && PROJECT_IMPORT_VALUE.trim()) {
    return PROJECT_IMPORT_VALUE.trim();
  }

  if (!notion || !PROJECTS_DB_ID) {
    return "Ametek SAP S4 Implementation Project";
  }

  if (PROJECT_PAGE_ID && PROJECT_PAGE_ID.trim()) {
    const page = await notion.pages.retrieve({ page_id: PROJECT_PAGE_ID.trim() });
    if (page.object === "page" && "properties" in page) {
      const title = (page.properties.Name?.title || [])
        .map((t) => t.plain_text)
        .join("")
        .trim();
      if (title) return title;
    }
  }

  const results = await notion.databases.query({
    database_id: PROJECTS_DB_ID,
    page_size: 100
  });

  for (const row of results.results) {
    if (!(row.object === "page" && "properties" in row)) continue;
    const title = (row.properties.Name?.title || [])
      .map((t) => t.plain_text)
      .join("")
      .trim();
    if (title.toLowerCase().includes("ametek")) {
      return title;
    }
  }

  return "Ametek SAP S4 Implementation Project";
}

(async function main() {
  ensureFolders();

  const sourceFile = pickPrimaryTaskCsv();
  const supplementalFile = findOptionalStagingFile(SUPPLEMENTAL_PLAN_NAME);

  const fileName = path.basename(sourceFile);
  const { raw, rows } = parseCsvRows(sourceFile);
  const supplementalByUid = parseSupplementalPlan(supplementalFile);

  for (const row of rows) {
    const uid = parseNumber(firstValue(row, "Unique ID"));
    if (uid == null) continue;
    const sup = supplementalByUid.get(uid);
    if (!sup) continue;

    // Task CSV is source of truth for owner names.
    // Never overwrite Resource Names / Business Validation Owner from supplemental files.
    if (sup.workstream) row[TASK_WORKSTREAM_FIELD] = sup.workstream;
    if (!firstValue(row, "Task Name") && sup.taskName) row["Task Name"] = sup.taskName;
  }

  const hash = crypto.createHash("sha1").update(raw).digest("hex").slice(0, 10);
  const stamp = new Date();
  const stampKey = stamp.toISOString().replace(/[:.]/g, "-");
  const importVersion = `${fileName}__${hash}__${stampKey}`;
  const importedAt = stamp.toISOString().slice(0, 10);
  const projectImportName = await resolveProjectImportValue();

  const headers = [
    "Task Name",
    "Unique ID",
    "WBS",
    "Outline Number",
    "Outline Level",
    "Summary",
    "Task Mode",
    "% Complete",
    "Duration",
    "Start",
    "Finish",
    "Actual Start",
    "Baseline10 Start",
    "Baseline10 Finish",
    "Baseline Start",
    "Baseline Finish",
    "Constraint Type",
    "Constraint Date",
    "Critical",
    "Milestone",
    "Total Slack",
    "Free Slack",
    "Predecessors",
    "Unique ID Predecessors",
    "Unique ID Successors",
    "WBS Predecessors",
    "WBS Successors",
    "Resource Names",
    "Assignee",
    "Business Validation Owner",
    "Workstream",
    "Resource Group",
    "Resource Type",
    "Resource Email (Text1)",
    "Resource Email (Text1) (Text1)",
    "Assignment",
    "Task Summary Name",
    "Number19",
    "Status",
    "Warning",
    "Error Message",
    "Status Update",
    "Active",
    "Placeholder",
    "Linked Fields",
    "Notes",

    "Parent Task",
    "Predecessor Tasks",
    "Successor Tasks",
    "Projects",
    "Assignee Team Members",
    "Business Validation Team Members",

    "Parent Unique ID (helper)",
    "Predecessor Unique IDs (helper)",
    "Successor Unique IDs (helper)",

    "MSP Source File",
    "MSP Import Version",
    "MSP Imported At",
    "MSP Row Hash",
    "MSP Relation Hash"
  ];

  const byHierarchy = new Map();
  for (const row of rows) {
    const h = firstValue(row, "Outline Number") || firstValue(row, "WBS");
    const uid = parseNumber(firstValue(row, "Unique ID"));
    if (h && uid != null) byHierarchy.set(h, uid);
  }

  const outRows = [];

  for (const row of rows) {
    const uniqueId = parseNumber(firstValue(row, "Unique ID"));
    if (uniqueId == null) continue;

    const outlineNumber = firstValue(row, "Outline Number");
    const wbs = firstValue(row, "WBS");
    const hierarchy = outlineNumber || wbs;
    const parentHierarchy = parentHierarchyKey(hierarchy);
    const parentUid = parentHierarchy ? (byHierarchy.get(parentHierarchy) || "") : "";

    const predIds = parseIdList(firstValue(row, "Unique ID Predecessors"));
    const succIds = parseIdList(firstValue(row, "Unique ID Successors"));

    const relHash = sha1Text(
      JSON.stringify({
        parentUniqueId: parentUid || "",
        predecessors: predIds,
        successors: succIds,
        project: projectImportName,
        assigneeNames: firstValue(row, "Resource Names"),
        bpoNames: firstValue(row, "Business Validation Owner")
      })
    );

    outRows.push({
      "Task Name": firstValue(row, "Task Name"),
      "Unique ID": uniqueId,
      WBS: wbs,
      "Outline Number": outlineNumber,
      "Outline Level": parseNumber(firstValue(row, "Outline Level")) ?? "",
      Summary: boolToText(parseBoolYesNo(firstValue(row, "Summary"))),
      "Task Mode": firstValue(row, "Task Mode"),
      "% Complete": parseNumber(firstValue(row, "% Complete")) ?? "",
      Duration: firstValue(row, "Duration"),
      Start: parseDate(firstValue(row, "Start")),
      Finish: parseDate(firstValue(row, "Finish")),
      "Actual Start": parseDate(firstValue(row, "Actual Start")),
      "Baseline10 Start": parseDate(firstValue(row, "Baseline10 Start")),
      "Baseline10 Finish": parseDate(firstValue(row, "Baseline10 Finish")),
      "Baseline Start": parseDate(firstValue(row, "Baseline Start")),
      "Baseline Finish": parseDate(firstValue(row, "Baseline Finish")),
      "Constraint Type": firstValue(row, "Constraint Type"),
      "Constraint Date": parseDate(firstValue(row, "Constraint Date")),
      Critical: boolToText(parseBoolYesNo(firstValue(row, "Critical"))),
      Milestone: boolToText(parseBoolYesNo(firstValue(row, "Milestone"))),
      "Total Slack": firstValue(row, "Total Slack"),
      "Free Slack": firstValue(row, "Free Slack"),
      Predecessors: firstValue(row, "Predecessors"),
      "Unique ID Predecessors": firstValue(row, "Unique ID Predecessors"),
      "Unique ID Successors": firstValue(row, "Unique ID Successors"),
      "WBS Predecessors": firstValue(row, "WBS Predecessors"),
      "WBS Successors": firstValue(row, "WBS Successors"),
      "Resource Names": firstValue(row, "Resource Names"),
      Assignee: firstValue(row, "Resource Names"),
      "Business Validation Owner": firstValue(row, "Business Validation Owner"),
      Workstream: firstValue(row, TASK_WORKSTREAM_FIELD),
      "Resource Group": firstValue(row, "Resource Group"),
      "Resource Type": firstValue(row, "Resource Type"),
      "Resource Email (Text1)": firstValue(row, "Resource Email (Text1) (Text1)"),
      "Resource Email (Text1) (Text1)": firstValue(row, "Resource Email (Text1) (Text1)"),
      Assignment: firstValue(row, "Assignment"),
      "Task Summary Name": firstValue(row, "Task Summary Name"),
      Number19: firstValue(row, "Number19"),
      Status: firstValue(row, "Status"),
      Warning: firstValue(row, "Warning"),
      "Error Message": firstValue(row, "Error Message"),
      "Status Update": deriveStatusUpdate(parseNumber(firstValue(row, "% Complete"))),
      Active: boolToText(parseBoolYesNo(firstValue(row, "Active"))),
      Placeholder: boolToText(parseBoolYesNo(firstValue(row, "Placeholder"))),
      "Linked Fields": boolToText(parseBoolYesNo(firstValue(row, "Linked Fields"))),
      Notes: NOTES_OWNER === "msp" ? firstValue(row, "Notes") : "",

      "Parent Task": "",
      "Predecessor Tasks": "",
      "Successor Tasks": "",
      Projects: projectImportName,
      "Assignee Team Members": firstValue(row, "Resource Names"),
      "Business Validation Team Members": firstValue(row, "Business Validation Owner"),

      "Parent Unique ID (helper)": parentUid,
      "Predecessor Unique IDs (helper)": predIds.join(","),
      "Successor Unique IDs (helper)": succIds.join(","),

      "MSP Source File": fileName,
      "MSP Import Version": importVersion,
      "MSP Imported At": importedAt,
      "MSP Row Hash": buildRowHash(row),
      "MSP Relation Hash": relHash
    });
  }

  const outName = `${stampKey}__manual-notion-import__${fileName}`;
  const outPath = path.join(EXPORTS_DIR, outName);
  writeCsv(outPath, headers, outRows);

  console.log("Manual import CSV generated:");
  console.log(outPath);
  console.log(`Rows: ${outRows.length}`);
  console.log(`Projects value used: ${projectImportName}`);
  if (supplementalFile) {
    console.log(`Merged supplemental fields from: ${supplementalFile}`);
  }
})();

/* eslint-disable no-console */
require("dotenv/config");

const fs = require("node:fs");
const path = require("node:path");
const crypto = require("node:crypto");
const { parse } = require("csv-parse/sync");
const { Client } = require("@notionhq/client");

const notion = new Client({ auth: process.env.NOTION_API_KEY });

const TASKS_DB_ID = process.env.NOTION_TASKS_DB_ID;
const PROJECTS_DB_ID = process.env.NOTION_PROJECTS_DB_ID;
const PROJECT_PAGE_ID = process.env.NOTION_AMETEK_PROJECT_PAGE_ID;
const TEAM_MEMBERS_DB_ID = process.env.NOTION_TEAM_MEMBERS_DB_ID;
const KEEP_LEGACY_TEXT_OWNER_FIELDS =
  (process.env.KEEP_LEGACY_TEXT_OWNER_FIELDS || "false").toLowerCase() === "true";

const NOTES_OWNER = (process.env.NOTES_OWNER || "notion").toLowerCase();
const IMPORT_CONCURRENCY = Number(process.env.IMPORT_CONCURRENCY || 1);
const REQUEST_DELAY_MS = Number(process.env.IMPORT_REQUEST_DELAY_MS || 250);
const API_TIMEOUT_MS = Number(process.env.IMPORT_API_TIMEOUT_MS || 45000);
const IMPORT_MODE =
  process.argv.includes("--changed-only") ||
  (process.env.IMPORT_MODE || "").toLowerCase() === "changed"
    ? "changed"
    : "full";

const ROOT = process.cwd();
const IMPORTS_DIR = path.join(ROOT, "imports");
const STAGING_DIR = path.join(IMPORTS_DIR, "staging");
const ARCHIVE_DIR = path.join(IMPORTS_DIR, "archive");
const LOGS_DIR = path.join(IMPORTS_DIR, "logs");
const LOGS_ARCHIVE_DIR = path.join(LOGS_DIR, "archive");

const SUPPLEMENTAL_PLAN_NAME = "Project plan with resources and busines.csv";
const RESOURCE_TABLE_NAME = "Resource table.csv";
const PROJECT_TEAM_NAME = "project team.csv";

const TASK_ASSIGNEE_TEAM_REL = "Assignee Team Members";
const TASK_BPO_TEAM_REL = "Business Validation Team Members";
const TASK_WORKSTREAM_FIELD = "Workstream";

const AUDIT_FIELD_SOURCE = "MSP Source File";
const AUDIT_FIELD_VERSION = "MSP Import Version";
const AUDIT_FIELD_AT = "MSP Imported At";
const AUDIT_FIELD_ROW_HASH = "MSP Row Hash";
const AUDIT_FIELD_REL_HASH = "MSP Relation Hash";

function assertEnv() {
  if (!process.env.NOTION_API_KEY) {
    throw new Error("NOTION_API_KEY is required in .env");
  }
  if (!TASKS_DB_ID) {
    throw new Error("NOTION_TASKS_DB_ID is required in .env");
  }
  if (!["msp", "notion"].includes(NOTES_OWNER)) {
    throw new Error('NOTES_OWNER must be either "msp" or "notion"');
  }
  if (!Number.isFinite(IMPORT_CONCURRENCY) || IMPORT_CONCURRENCY < 1) {
    throw new Error("IMPORT_CONCURRENCY must be a positive integer.");
  }
  if (!Number.isFinite(REQUEST_DELAY_MS) || REQUEST_DELAY_MS < 0) {
    throw new Error("IMPORT_REQUEST_DELAY_MS must be 0 or greater.");
  }
  if (!Number.isFinite(API_TIMEOUT_MS) || API_TIMEOUT_MS < 1000) {
    throw new Error("IMPORT_API_TIMEOUT_MS must be at least 1000.");
  }
}

function sleep(ms) {
  return new Promise((resolve) => setTimeout(resolve, ms));
}

function normalizeName(value) {
  return String(value || "")
    .trim()
    .toLowerCase()
    .replace(/\s+/g, " ");
}

function splitNameList(value) {
  if (!value) return [];
  return [...new Set(
    String(value)
      .split(/[;|]/)
      .map((v) => v.trim())
      .filter(Boolean)
  )];
}

function isRetryableError(error) {
  const status = error?.status || error?.body?.status;
  const code = error?.code || error?.body?.code;
  return (
    status === 429 ||
    status >= 500 ||
    code === "rate_limited" ||
    code === "request_timeout"
  );
}

async function withTimeout(promise, ms, label) {
  let timeoutHandle;
  const timeoutPromise = new Promise((_, reject) => {
    timeoutHandle = setTimeout(() => {
      const err = new Error(`Timeout after ${ms}ms: ${label}`);
      err.code = "request_timeout";
      reject(err);
    }, ms);
  });

  try {
    return await Promise.race([promise, timeoutPromise]);
  } finally {
    clearTimeout(timeoutHandle);
  }
}

async function withRetry(fn, label, maxRetries = 5) {
  let attempt = 0;

  while (true) {
    try {
      return await withTimeout(fn(), API_TIMEOUT_MS, label);
    } catch (error) {
      attempt += 1;
      if (!isRetryableError(error) || attempt > maxRetries) {
        throw error;
      }

      const retryAfterSec = Number(
        error?.headers?.get?.("retry-after") || error?.body?.retry_after
      );
      const retryAfterMs = Number.isFinite(retryAfterSec) ? retryAfterSec * 1000 : 0;
      const delayMs = Math.max(retryAfterMs, Math.min(2500 * attempt, 15000));
      console.log(
        `Retrying ${label} after ${delayMs}ms (attempt ${attempt}/${maxRetries})`
      );
      await sleep(delayMs);
    }
  }
}

async function runWithConcurrency(items, limit, handler, phaseLabel) {
  const startedAt = Date.now();
  let cursor = 0;
  let completed = 0;
  let nextLogAt = 200;

  async function worker() {
    while (true) {
      const index = cursor;
      if (index >= items.length) {
        return;
      }
      cursor += 1;

      if (REQUEST_DELAY_MS > 0) {
        await sleep(REQUEST_DELAY_MS);
      }

      await handler(items[index], index);
      completed += 1;

      if (completed >= nextLogAt || completed === items.length) {
        const seconds = Math.round((Date.now() - startedAt) / 1000);
        console.log(`${phaseLabel}: ${completed}/${items.length} completed (${seconds}s)`);
        nextLogAt += 200;
      }
    }
  }

  const workers = Array.from({ length: Math.min(limit, items.length) }, () => worker());
  await Promise.all(workers);
}

function ensureFolders() {
  fs.mkdirSync(STAGING_DIR, { recursive: true });
  fs.mkdirSync(ARCHIVE_DIR, { recursive: true });
  fs.mkdirSync(LOGS_DIR, { recursive: true });
  fs.mkdirSync(LOGS_ARCHIVE_DIR, { recursive: true });
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

function parseDateValue(value) {
  if (!value) return null;
  const text = String(value).trim();
  if (!text) return null;
  const d = new Date(text);
  if (Number.isNaN(d.getTime())) return null;
  return d.toISOString().slice(0, 10);
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

function parseResourceList(value) {
  if (!value) return [];
  return String(value)
    .split(/[;,|]/)
    .map((s) => s.trim())
    .filter(Boolean);
}

function canonicalHeader(value) {
  return String(value || "")
    .replace(/\s+/g, " ")
    .trim();
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

function parseResourceTable(filePath) {
  const result = {
    headers: [],
    rowsByName: new Map()
  };

  if (!filePath || !fs.existsSync(filePath)) {
    return result;
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
    (row) => Array.isArray(row) && row.includes("Resource name") && row.includes("Abbreviation")
  );
  if (headerIndex === -1) {
    return result;
  }

  const header = matrix[headerIndex].map((h) => String(h || "").trim());
  result.headers = header.filter(Boolean);

  const idxByName = new Map();
  header.forEach((h, i) => {
    idxByName.set(h, i);
  });

  for (let i = headerIndex + 1; i < matrix.length; i += 1) {
    const row = matrix[i] || [];
    const name = String(row[idxByName.get("Resource name")] || "").trim();
    if (!name) continue;

    const key = normalizeName(name);
    const existing = result.rowsByName.get(key) || {
      name,
      values: {},
      workstreams: new Set()
    };

    for (const h of result.headers) {
      const idx = idxByName.get(h);
      const value = String(row[idx] || "").trim();
      if (!value) continue;

      if (h === "Workstream") {
        existing.workstreams.add(value);
      } else if (!existing.values[h]) {
        existing.values[h] = value;
      }
    }

    existing.name = existing.name || name;
    result.rowsByName.set(key, existing);
  }

  return result;
}

function parseProjectTeamRoster(filePath) {
  const result = {
    headers: [],
    rowsByName: new Map()
  };

  if (!filePath || !fs.existsSync(filePath)) {
    return result;
  }

  const raw = fs.readFileSync(filePath, "utf8");
  const matrix = parse(raw, {
    columns: false,
    bom: true,
    skip_empty_lines: false,
    trim: true,
    relax_column_count: true
  });

  const headerIndex = matrix.findIndex((row) => {
    const cells = (row || []).map(canonicalHeader);
    return (
      cells.includes("Name") &&
      cells.some((c) => c.toLowerCase().includes("email")) &&
      cells.some((c) => c.toLowerCase().includes("project position"))
    );
  });

  if (headerIndex === -1) {
    return result;
  }

  const header = (matrix[headerIndex] || []).map(canonicalHeader);
  const idxByHeader = new Map();
  header.forEach((h, i) => {
    if (h) idxByHeader.set(h, i);
  });

  result.headers = [...idxByHeader.keys()];

  const nameHeader = "Name";
  const nameIdx = idxByHeader.get(nameHeader);
  if (nameIdx == null) {
    return result;
  }

  for (let i = headerIndex + 1; i < matrix.length; i += 1) {
    const row = matrix[i] || [];
    const name = canonicalHeader(row[nameIdx]);
    if (!name) continue;
    if (name.toLowerCase().includes("if not assign")) continue;

    const key = normalizeName(name);
    if (!key) continue;

    const values = {};
    for (const h of result.headers) {
      const idx = idxByHeader.get(h);
      const value = canonicalHeader(row[idx]);
      if (value) values[h] = value;
    }

    result.rowsByName.set(key, {
      name,
      values
    });
  }

  return result;
}

function normalizePhrase(value) {
  return String(value || "")
    .toLowerCase()
    .replace(/[^a-z0-9\s/&-]/g, " ")
    .replace(/\s+/g, " ")
    .trim();
}

function containsPhrase(haystack, needle) {
  const h = normalizePhrase(haystack);
  const n = normalizePhrase(needle);
  if (!h || !n) return false;

  if (n.length <= 3) {
    const re = new RegExp(`\\b${n.replace(/[.*+?^${}()|[\\]\\]/g, "\\$&")}\\b`, "i");
    return re.test(h);
  }

  return h.includes(n);
}

function splitDelimited(value) {
  return String(value || "")
    .split(/[;,|/]/)
    .map((x) => x.trim())
    .filter(Boolean);
}

function buildAssignmentProfiles(projectTeam, resourceTable) {
  const profiles = new Map();

  function ensureProfile(name) {
    const key = normalizeName(name);
    if (!key) return null;
    if (!profiles.has(key)) {
      profiles.set(key, {
        name: String(name || "").trim(),
        active: true,
        workstreams: new Set(),
        assignmentTerms: new Set(),
        positionTerms: new Set(),
        isBpoLike: false
      });
    }
    return profiles.get(key);
  }

  for (const row of projectTeam.rowsByName.values()) {
    const p = ensureProfile(row.name);
    if (!p) continue;

    const status = String(row.values?.Status || "").trim().toLowerCase();
    if (status && status !== "active") p.active = false;

    const assignment = String(row.values?.["Project Team Assignment"] || "").trim();
    for (const term of splitDelimited(assignment)) {
      p.assignmentTerms.add(term);
      p.workstreams.add(term);
    }

    const position = String(row.values?.["Project Position"] || "").trim();
    for (const term of splitDelimited(position)) {
      p.positionTerms.add(term);
    }

    if (/\bbpo\b|business process owner/i.test(`${assignment} ${position}`)) {
      p.isBpoLike = true;
    }
  }

  for (const row of resourceTable.rowsByName.values()) {
    const p = ensureProfile(row.name);
    if (!p) continue;

    for (const ws of row.workstreams || []) {
      if (ws) {
        p.workstreams.add(String(ws).trim());
      }
    }

    if (String(row.values?.BPO || "").trim()) {
      p.isBpoLike = true;
    }
  }

  return [...profiles.values()].filter((p) => p.active !== false);
}

function isWorkingTaskRow(row) {
  const isSummary = parseBoolYesNo(firstValue(row, "Summary"));
  const isMilestone = parseBoolYesNo(firstValue(row, "Milestone"));
  const pct = parseNumber(firstValue(row, "% Complete"));
  const isDone = pct != null && pct >= 100;
  const isActiveRaw = firstValue(row, "Active");
  const isActive = !isActiveRaw || parseBoolYesNo(isActiveRaw);
  const isPlaceholder = parseBoolYesNo(firstValue(row, "Placeholder"));
  return !isSummary && !isMilestone && !isDone && isActive && !isPlaceholder;
}

function pickAssigneesFromDescription(taskText, profiles) {
  const scored = [];

  for (const p of profiles) {
    let score = 0;

    if (containsPhrase(taskText, p.name)) score += 12;

    const nameParts = p.name.split(/\s+/).filter((x) => x.length >= 4);
    for (const np of nameParts) {
      if (containsPhrase(taskText, np)) score += 4;
    }

    for (const term of p.workstreams) {
      if (containsPhrase(taskText, term)) score += 6;
    }
    for (const term of p.assignmentTerms) {
      if (containsPhrase(taskText, term)) score += 5;
    }
    for (const term of p.positionTerms) {
      if (containsPhrase(taskText, term)) score += 2;
    }

    if (score > 0) {
      scored.push({ profile: p, score });
    }
  }

  scored.sort((a, b) => b.score - a.score || a.profile.name.localeCompare(b.profile.name));
  return scored;
}

function enrichWorkingTasksFromDescription(rows, projectTeam, resourceTable) {
  const profiles = buildAssignmentProfiles(projectTeam, resourceTable);
  if (!profiles.length) {
    return {
      profiles: 0,
      rowsConsidered: 0,
      resourceRowsUpdated: 0,
      bpoRowsUpdated: 0,
      workstreamRowsUpdated: 0
    };
  }

  let rowsConsidered = 0;
  let resourceRowsUpdated = 0;
  let bpoRowsUpdated = 0;
  let workstreamRowsUpdated = 0;

  for (const row of rows) {
    if (!isWorkingTaskRow(row)) continue;

    rowsConsidered += 1;
    const taskText = `${firstValue(row, "Task Name")} ${firstValue(row, "Task description")}`.trim();
    if (!taskText) continue;

    const matches = pickAssigneesFromDescription(taskText, profiles).slice(0, 6);
    if (!matches.length) continue;

    const matchedWorkstreams = [...new Set(
      matches.flatMap((m) => [...m.profile.workstreams]).filter(Boolean)
    )];

    const resourceNames = [...new Set(
      matches
        .filter((m) => !m.profile.isBpoLike)
        .map((m) => m.profile.name)
        .filter(Boolean)
    )];

    const bpoNames = [...new Set(
      matches
        .filter((m) => m.profile.isBpoLike)
        .map((m) => m.profile.name)
        .filter(Boolean)
    )];

    if (!firstValue(row, "Resource Names") && resourceNames.length) {
      row["Resource Names"] = resourceNames.slice(0, 3).join("; ");
      resourceRowsUpdated += 1;
    }

    if (!firstValue(row, "Business Validation Owner") && bpoNames.length) {
      row["Business Validation Owner"] = bpoNames.slice(0, 3).join("; ");
      bpoRowsUpdated += 1;
    }

    if (!firstValue(row, TASK_WORKSTREAM_FIELD) && matchedWorkstreams.length) {
      row[TASK_WORKSTREAM_FIELD] = matchedWorkstreams.slice(0, 2).join("; ");
      workstreamRowsUpdated += 1;
    }
  }

  return {
    profiles: profiles.length,
    rowsConsidered,
    resourceRowsUpdated,
    bpoRowsUpdated,
    workstreamRowsUpdated
  };
}

function mergeSupplementalIntoBaseRows(baseRows, supplementalByUid) {
  for (const row of baseRows) {
    const uid = parseNumber(firstValue(row, "Unique ID"));
    if (uid == null) continue;
    const sup = supplementalByUid.get(uid);
    if (!sup) continue;

    // Task CSV is source of truth for owner names.
    // Never overwrite Resource Names / Business Validation Owner from supplemental files.
    if (sup.workstream) row[TASK_WORKSTREAM_FIELD] = sup.workstream;
    if (!firstValue(row, "Task Name") && sup.taskName) {
      row["Task Name"] = sup.taskName;
    }
  }
}

function collectTeamSignals(rows, resourceTable) {
  const byName = new Map();

  function upsertSignal(name) {
    const key = normalizeName(name);
    if (!key) return null;
    if (!byName.has(key)) {
      byName.set(key, {
        name,
        inResourceField: false,
        inBusinessValidationField: false,
        workstreams: new Set()
      });
    }
    return byName.get(key);
  }

  for (const row of rows) {
    const rowWorkstream = firstValue(row, TASK_WORKSTREAM_FIELD);

    for (const name of splitNameList(firstValue(row, "Resource Names"))) {
      const signal = upsertSignal(name);
      if (!signal) continue;
      signal.inResourceField = true;
      if (rowWorkstream) signal.workstreams.add(rowWorkstream);
    }

    for (const name of splitNameList(firstValue(row, "Business Validation Owner"))) {
      const signal = upsertSignal(name);
      if (!signal) continue;
      signal.inBusinessValidationField = true;
      if (rowWorkstream) signal.workstreams.add(rowWorkstream);
    }
  }

  // Do not seed names from lookup tables; tasks are the source of truth.

  return byName;
}

function rosterHeaderToTeamProperty(header) {
  const normalized = canonicalHeader(header);
  if (!normalized) return null;

  if (normalized === "Email (AMETEK)") return "Email Address";
  if (normalized === "Contact Number") return "Phone Number";
  if (normalized === "Project Team Assignment") return "Project Team Assignment";
  if (normalized === "Start Date") return "Start Date";
  if (normalized === "Status") return "Status";

  return normalized;
}

function listDuplicateHeaders(rawCsvText) {
  const headerRaw = rawCsvText.split(/\r?\n/)[0] || "";
  const headers = headerRaw.split(",").map((h) => h.trim());
  const counts = {};
  for (const h of headers) {
    counts[h] = (counts[h] || 0) + 1;
  }
  return Object.entries(counts)
    .filter(([, count]) => count > 1)
    .map(([name, count]) => ({ name, count }));
}

function getHierarchyKey(row) {
  return firstValue(row, "Outline Number") || firstValue(row, "WBS") || "";
}

function parentHierarchyKey(key) {
  if (!key) return "";
  const parts = key.split(".").map((x) => x.trim()).filter(Boolean);
  if (parts.length <= 1) return "";
  return parts.slice(0, -1).join(".");
}

function analyzeCsvPreflight(rows, duplicateHeaders) {
  const coreHeaders = [
    "Unique ID",
    "Task Name",
    "WBS",
    "Outline Number",
    "Outline Level",
    "Start",
    "Finish",
    "% Complete",
    "Critical",
    "Milestone",
    "Unique ID Predecessors",
    "Unique ID Successors"
  ];

  const missingCoreHeaders = coreHeaders.filter(
    (header) => !rows.length || !(header in rows[0])
  );

  let uidMissing = 0;
  let uidDuplicate = 0;
  let hierarchyMissing = 0;
  let hierarchyDuplicate = 0;
  let percentInvalid = 0;
  let startInvalid = 0;
  let finishInvalid = 0;

  let tasksWithPredecessors = 0;
  let tasksWithSuccessors = 0;
  let predecessorRefs = 0;
  let successorRefs = 0;
  let unresolvedPredecessorRefs = 0;
  let unresolvedSuccessorRefs = 0;

  const uidSet = new Set();
  const hierarchySet = new Set();
  const byUid = new Map();

  for (const row of rows) {
    const uid = parseNumber(firstValue(row, "Unique ID"));
    if (uid == null) {
      uidMissing += 1;
    } else {
      if (uidSet.has(uid)) uidDuplicate += 1;
      uidSet.add(uid);
      byUid.set(uid, row);
    }

    const hierarchy = getHierarchyKey(row);
    if (!hierarchy) {
      hierarchyMissing += 1;
    } else {
      if (hierarchySet.has(hierarchy)) hierarchyDuplicate += 1;
      hierarchySet.add(hierarchy);
    }

    const pct = firstValue(row, "% Complete");
    if (pct) {
      const num = parseNumber(pct);
      if (num == null || num < 0 || num > 100) percentInvalid += 1;
    }

    const start = firstValue(row, "Start");
    if (start && Number.isNaN(new Date(start).getTime())) startInvalid += 1;

    const finish = firstValue(row, "Finish");
    if (finish && Number.isNaN(new Date(finish).getTime())) finishInvalid += 1;
  }

  for (const row of byUid.values()) {
    const preds = parseIdList(firstValue(row, "Unique ID Predecessors"));
    const succs = parseIdList(firstValue(row, "Unique ID Successors"));

    if (preds.length > 0) tasksWithPredecessors += 1;
    if (succs.length > 0) tasksWithSuccessors += 1;

    predecessorRefs += preds.length;
    successorRefs += succs.length;

    for (const pred of preds) {
      if (!uidSet.has(pred)) unresolvedPredecessorRefs += 1;
    }
    for (const succ of succs) {
      if (!uidSet.has(succ)) unresolvedSuccessorRefs += 1;
    }
  }

  return {
    rowCount: rows.length,
    duplicateHeaders,
    missingCoreHeaders,
    uniqueId: {
      missingRows: uidMissing,
      duplicateRows: uidDuplicate
    },
    hierarchy: {
      missingRows: hierarchyMissing,
      duplicateRows: hierarchyDuplicate
    },
    validation: {
      percentInvalid,
      startInvalid,
      finishInvalid
    },
    dependency: {
      tasksWithPredecessors,
      tasksWithoutPredecessors: byUid.size - tasksWithPredecessors,
      tasksWithSuccessors,
      tasksWithoutSuccessors: byUid.size - tasksWithSuccessors,
      predecessorRefs,
      unresolvedPredecessorRefs,
      successorRefs,
      unresolvedSuccessorRefs
    },
    decision:
      missingCoreHeaders.length === 0 &&
      duplicateHeaders.length === 0 &&
      uidMissing === 0 &&
      uidDuplicate === 0 &&
      hierarchyMissing === 0 &&
      hierarchyDuplicate === 0 &&
      percentInvalid === 0 &&
      startInvalid === 0 &&
      finishInvalid === 0
        ? "GO"
        : "NO-GO"
  };
}

function writeImportReport(report) {
  const latestPath = path.join(LOGS_DIR, "latest-import-report.json");
  const previousLatestExists = fs.existsSync(latestPath);

  if (previousLatestExists) {
    const previous = JSON.parse(fs.readFileSync(latestPath, "utf8"));
    const previousStamp =
      previous?.runtime?.stampKey ||
      new Date().toISOString().replace(/[:.]/g, "-");
    const archivePath = path.join(
      LOGS_ARCHIVE_DIR,
      `${previousStamp}__import-report.json`
    );
    fs.renameSync(latestPath, archivePath);
  }

  fs.writeFileSync(latestPath, JSON.stringify(report, null, 2), "utf8");
}

function dateProperty(value) {
  const parsed = parseDateValue(value);
  if (!parsed) {
    return { date: null };
  }
  return { date: { start: parsed } };
}

function richText(value) {
  if (!value) return [];
  return [{ type: "text", text: { content: String(value) } }];
}

function titleText(value) {
  const text = String(value || "").trim() || "(untitled task)";
  return [{ type: "text", text: { content: text.slice(0, 1900) } }];
}

function sha1Text(value) {
  return crypto.createHash("sha1").update(value).digest("hex");
}

function relationArray(pageIds) {
  return pageIds.map((id) => ({ id }));
}

function deriveStatusUpdate(percentComplete) {
  const pct = Number(percentComplete);
  if (!Number.isFinite(pct)) return "Not started";
  if (pct >= 100) return "Done";
  if (pct >= 25) return "In progress";
  return "Not started";
}

function filterPropertiesBySchema(properties, schemaProperties) {
  const filtered = {};
  for (const [name, value] of Object.entries(properties)) {
    if (schemaProperties[name]) filtered[name] = value;
  }
  return filtered;
}

async function ensureAuditProperties() {
  const payload = {
    database_id: TASKS_DB_ID,
    properties: {
      [AUDIT_FIELD_SOURCE]: { rich_text: {} },
      [AUDIT_FIELD_VERSION]: { rich_text: {} },
      [AUDIT_FIELD_AT]: { date: {} },
      [AUDIT_FIELD_ROW_HASH]: { rich_text: {} },
      [AUDIT_FIELD_REL_HASH]: { rich_text: {} }
    }
  };
  await withRetry(() => notion.databases.update(payload), "ensure-audit-properties");
}

async function fetchAllDatabasePages(databaseId, label) {
  const pages = [];
  let cursor = undefined;

  do {
    const res = await withRetry(
      () =>
        notion.databases.query({
          database_id: databaseId,
          start_cursor: cursor,
          page_size: 100
        }),
      `fetch-${label}`
    );

    for (const r of res.results) {
      if (r.object === "page" && "properties" in r) pages.push(r);
    }

    cursor = res.has_more ? res.next_cursor : undefined;
  } while (cursor);

  return pages;
}

function getNumberProperty(page, propName) {
  const prop = page.properties[propName];
  if (!prop || prop.type !== "number") return null;
  return prop.number;
}

function getRichTextProperty(page, propName) {
  const prop = page.properties[propName];
  if (!prop || prop.type !== "rich_text") return "";
  return prop.rich_text.map((x) => x.plain_text).join("");
}

function getTitleProperty(page, propName = "Name") {
  const prop = page.properties[propName];
  if (!prop || prop.type !== "title") return "";
  return prop.title.map((x) => x.plain_text).join("");
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

function buildRelationHash({
  parentId,
  predecessorIds,
  successorIds,
  projectPageId,
  assigneeTeamMemberIds,
  bpoTeamMemberIds
}) {
  const payload = {
    parentId: parentId || "",
    predecessorIds: [...new Set(predecessorIds)].sort(),
    successorIds: [...new Set(successorIds)].sort(),
    projectPageId: projectPageId || "",
    assigneeTeamMemberIds: [...new Set(assigneeTeamMemberIds || [])].sort(),
    bpoTeamMemberIds: [...new Set(bpoTeamMemberIds || [])].sort()
  };
  return sha1Text(JSON.stringify(payload));
}

async function findAmetekProjectPageId() {
  if (PROJECT_PAGE_ID && PROJECT_PAGE_ID.trim()) {
    return PROJECT_PAGE_ID.trim();
  }
  if (!PROJECTS_DB_ID) {
    return undefined;
  }

  const res = await withRetry(
    () =>
      notion.databases.query({
        database_id: PROJECTS_DB_ID,
        page_size: 100
      }),
    "find-ametek-project"
  );

  for (const row of res.results) {
    if (!(row.object === "page" && "properties" in row)) continue;
    const nameProp = row.properties.Name;
    if (!nameProp || nameProp.type !== "title") continue;
    const text = nameProp.title.map((x) => x.plain_text).join("").toLowerCase();
    if (text.includes("ametek")) {
      return row.id;
    }
  }

  return undefined;
}

async function findTeamMembersDbId() {
  if (TEAM_MEMBERS_DB_ID && TEAM_MEMBERS_DB_ID.trim()) {
    return TEAM_MEMBERS_DB_ID.trim();
  }

  const res = await withRetry(
    () =>
      notion.search({
        query: "Team Member",
        filter: { value: "database", property: "object" },
        page_size: 100
      }),
    "find-team-members-db"
  );

  const db = res.results.find((r) => {
    if (r.object !== "database") return false;
    const title = (r.title || []).map((t) => t.plain_text).join("").toLowerCase();
    return title === "team member" || title === "team members";
  });

  if (!db) {
    throw new Error(
      "Could not find Team Member(s) database. Set NOTION_TEAM_MEMBERS_DB_ID in .env."
    );
  }

  return db.id;
}

async function ensureTaskEnhancementProperties(taskDbProperties, teamMembersDbId) {
  const addProps = {};

  if (!taskDbProperties[TASK_WORKSTREAM_FIELD]) {
    addProps[TASK_WORKSTREAM_FIELD] = { rich_text: {} };
  }

  if (!taskDbProperties[TASK_ASSIGNEE_TEAM_REL]) {
    addProps[TASK_ASSIGNEE_TEAM_REL] = {
      relation: { database_id: teamMembersDbId, single_property: {} }
    };
  }

  if (!taskDbProperties[TASK_BPO_TEAM_REL]) {
    addProps[TASK_BPO_TEAM_REL] = {
      relation: { database_id: teamMembersDbId, single_property: {} }
    };
  }

  if (Object.keys(addProps).length > 0) {
    await withRetry(
      () => notion.databases.update({ database_id: TASKS_DB_ID, properties: addProps }),
      "ensure-task-enhancement-properties"
    );
  }
}

function teamMemberSchemaFromSourceHeaders(resourceHeaders, rosterHeaders) {
  const schema = {
    Type: {
      select: {
        options: [{ name: "BPO" }, { name: "Func/Tech" }]
      }
    },
    Workstream: { rich_text: {} }
  };

  for (const h of resourceHeaders) {
    if (!h || h === "Resource name") continue;

    let propName = h;
    if (h === "Type") propName = "Source Type";

    if (schema[propName]) continue;

    if (["Rate", "Count", "Cost"].includes(h)) {
      schema[propName] = { number: {} };
    } else {
      schema[propName] = { rich_text: {} };
    }
  }

  for (const h of rosterHeaders) {
    const propName = rosterHeaderToTeamProperty(h);
    if (!propName || propName === "Name") continue;
    if (schema[propName]) continue;

    if (propName === "Email Address") {
      schema[propName] = { email: {} };
    } else if (propName === "Phone Number") {
      schema[propName] = { phone_number: {} };
    } else if (propName === "Start Date") {
      schema[propName] = { date: {} };
    } else if (propName === "Status") {
      schema[propName] = {
        status: {
          options: [{ name: "Active" }, { name: "Inactive" }]
        }
      };
    } else {
      schema[propName] = { rich_text: {} };
    }
  }

  return schema;
}

function resolveTeamStatusName(rawValue, statusOptions) {
  const options = (statusOptions || []).map((o) => String(o?.name || "").trim()).filter(Boolean);
  if (!options.length) return null;

  const raw = String(rawValue || "").trim().toLowerCase();
  if (!raw) return options[0];

  const exact = options.find((name) => name.toLowerCase() === raw);
  if (exact) return exact;

  const wantsInactive = /(inactive|roll.?off|offboard|left|former|closed|disabled)/i.test(raw);

  if (wantsInactive) {
    const inactiveLike = options.find((name) => /(inactive|off|roll|left|former|closed|disabled)/i.test(name));
    if (inactiveLike) return inactiveLike;
  } else {
    const activeLike = options.find((name) => /(active|current|enabled|working|open)/i.test(name));
    if (activeLike) return activeLike;
  }

  return options[0];
}

async function ensureTeamMemberProperties(teamMembersDbId, resourceHeaders, rosterHeaders) {
  const db = await withRetry(
    () => notion.databases.retrieve({ database_id: teamMembersDbId }),
    "retrieve-team-members-db"
  );

  const existing = db.properties || {};
  const wanted = teamMemberSchemaFromSourceHeaders(resourceHeaders, rosterHeaders);
  const addProps = {};

  for (const [name, spec] of Object.entries(wanted)) {
    if (!existing[name]) {
      addProps[name] = spec;
    }
  }

  if (Object.keys(addProps).length > 0) {
    await withRetry(
      () =>
        notion.databases.update({
          database_id: teamMembersDbId,
          properties: addProps
        }),
      "ensure-team-member-properties"
    );
  }
}

async function fetchWorkspaceUsersByEmail() {
  const usersByEmail = new Map();
  let cursor = undefined;

  do {
    const res = await withRetry(
      () => notion.users.list({ start_cursor: cursor, page_size: 100 }),
      "list-workspace-users"
    );

    for (const user of res.results || []) {
      const email = user?.person?.email;
      if (!email) continue;
      usersByEmail.set(String(email).toLowerCase(), user.id);
    }

    cursor = res.has_more ? res.next_cursor : undefined;
  } while (cursor);

  return usersByEmail;
}

function buildTeamMemberProperties({
  name,
  signal,
  resourceData,
  rosterData,
  projectPageId,
  teamDbProperties
}) {
  const props = {
    Name: { title: titleText(name) }
  };

  const typeValue = signal.inBusinessValidationField
    ? "BPO"
    : signal.inResourceField
      ? "Func/Tech"
      : "Func/Tech";
  props.Type = { select: { name: typeValue } };

  const workstreams = new Set(signal.workstreams || []);
  for (const ws of resourceData?.workstreams || []) workstreams.add(ws);
  props.Workstream = { rich_text: richText([...workstreams].join("; ")) };

  if (projectPageId) {
    props.Projects = { relation: relationArray([projectPageId]) };
  }

  const values = resourceData?.values || {};
  for (const [key, rawValue] of Object.entries(values)) {
    if (key === "Resource name") continue;

    const propName = key === "Type" ? "Source Type" : key;
    const schemaType = teamDbProperties[propName]?.type;

    if (!schemaType) continue;

    if (schemaType === "number") {
      props[propName] = { number: parseNumber(rawValue) };
    } else if (schemaType === "rich_text") {
      props[propName] = { rich_text: richText(rawValue) };
    } else if (schemaType === "email") {
      props[propName] = { email: rawValue || null };
    }
  }

  const rosterValues = rosterData?.values || {};
  for (const [sourceHeader, rawValue] of Object.entries(rosterValues)) {
    const propName = rosterHeaderToTeamProperty(sourceHeader);
    if (!propName || propName === "Name") continue;
    const schemaType = teamDbProperties[propName]?.type;
    if (!schemaType) continue;

    if (schemaType === "email") {
      props[propName] = { email: rawValue || null };
    } else if (schemaType === "phone_number") {
      props[propName] = { phone_number: rawValue || null };
    } else if (schemaType === "date") {
      const parsed = parseDateValue(rawValue);
      props[propName] = parsed ? { date: { start: parsed } } : { date: null };
    } else if (schemaType === "status") {
      const statusName = resolveTeamStatusName(rawValue, teamDbProperties[propName]?.status?.options || []);
      if (statusName) {
        props[propName] = { status: { name: statusName } };
      }
    } else if (schemaType === "rich_text") {
      props[propName] = { rich_text: richText(rawValue) };
    }
  }

  return filterPropertiesBySchema(props, teamDbProperties);
}

async function syncTeamMembers({
  teamMembersDbId,
  teamSignals,
  resourceTable,
  projectTeam,
  projectPageId
}) {
  await ensureTeamMemberProperties(
    teamMembersDbId,
    resourceTable.headers || [],
    projectTeam.headers || []
  );

  const teamDb = await withRetry(
    () => notion.databases.retrieve({ database_id: teamMembersDbId }),
    "retrieve-team-members-db-after-ensure"
  );
  const teamDbProperties = teamDb.properties || {};

  const existingPages = await fetchAllDatabasePages(teamMembersDbId, "team-members-pages");
  const existingByName = new Map();
  for (const p of existingPages) {
    const name = getTitleProperty(p, "Name");
    const key = normalizeName(name);
    if (key) existingByName.set(key, p);
  }

  const allNames = new Map();
  for (const [k, s] of teamSignals.entries()) allNames.set(k, s);

  let created = 0;
  let updated = 0;

  const teamMemberIdByName = new Map();
  const items = Array.from(allNames.entries());

  await runWithConcurrency(
    items,
    IMPORT_CONCURRENCY,
    async ([key, signal]) => {
      const name = signal.name;
      if (!name) return;

      const resourceData = resourceTable.rowsByName.get(key);
      const rosterData = projectTeam.rowsByName.get(key);
      const properties = buildTeamMemberProperties({
        name,
        signal,
        resourceData,
        rosterData,
        projectPageId,
        teamDbProperties
      });

      const existing = existingByName.get(key);
      if (existing) {
        await withRetry(
          () => notion.pages.update({ page_id: existing.id, properties }),
          `team-member-update-${name}`
        );
        teamMemberIdByName.set(key, existing.id);
        updated += 1;
      } else {
        const createdPage = await withRetry(
          () =>
            notion.pages.create({
              parent: { database_id: teamMembersDbId },
              properties
            }),
          `team-member-create-${name}`
        );
        teamMemberIdByName.set(key, createdPage.id);
        created += 1;
      }
    },
    "Team Members sync"
  );

  return {
    teamMemberIdByName,
    created,
    updated,
    total: items.length
  };
}

function buildMspProperties(row, audit, rowHash, schemaContext) {
  const resourceNames = firstValue(row, "Resource Names");
  const businessValidationOwner = firstValue(row, "Business Validation Owner");
  const percentComplete = parseNumber(firstValue(row, "% Complete"));

  const properties = {
    "Task Name": { title: titleText(firstValue(row, "Task Name")) },
    WBS: { rich_text: richText(firstValue(row, "WBS")) },
    "Outline Number": { rich_text: richText(firstValue(row, "Outline Number")) },
    "Outline Level": { number: parseNumber(firstValue(row, "Outline Level")) },
    Summary: { checkbox: parseBoolYesNo(firstValue(row, "Summary")) },
    "Task Mode": { rich_text: richText(firstValue(row, "Task Mode")) },
    "% Complete": { number: percentComplete },
    Duration: { rich_text: richText(firstValue(row, "Duration")) },
    Start: dateProperty(firstValue(row, "Start")),
    Finish: dateProperty(firstValue(row, "Finish")),
    "Actual Start": dateProperty(firstValue(row, "Actual Start")),
    "Baseline10 Start": dateProperty(firstValue(row, "Baseline10 Start")),
    "Baseline10 Finish": dateProperty(firstValue(row, "Baseline10 Finish")),
    "Baseline Start": dateProperty(firstValue(row, "Baseline Start")),
    "Baseline Finish": dateProperty(firstValue(row, "Baseline Finish")),
    "Constraint Type": { rich_text: richText(firstValue(row, "Constraint Type")) },
    "Constraint Date": dateProperty(firstValue(row, "Constraint Date")),
    Critical: { checkbox: parseBoolYesNo(firstValue(row, "Critical")) },
    Milestone: { checkbox: parseBoolYesNo(firstValue(row, "Milestone")) },
    "Total Slack": { rich_text: richText(firstValue(row, "Total Slack")) },
    "Free Slack": { rich_text: richText(firstValue(row, "Free Slack")) },
    Predecessors: { rich_text: richText(firstValue(row, "Predecessors")) },
    "Unique ID Predecessors": {
      rich_text: richText(firstValue(row, "Unique ID Predecessors"))
    },
    "Unique ID Successors": {
      rich_text: richText(firstValue(row, "Unique ID Successors"))
    },
    "WBS Predecessors": { rich_text: richText(firstValue(row, "WBS Predecessors")) },
    "WBS Successors": { rich_text: richText(firstValue(row, "WBS Successors")) },
    "Resource Names": {
      rich_text: KEEP_LEGACY_TEXT_OWNER_FIELDS ? richText(resourceNames) : []
    },
    "Business Validation Owner": {
      rich_text: KEEP_LEGACY_TEXT_OWNER_FIELDS ? richText(businessValidationOwner) : []
    },
    [TASK_WORKSTREAM_FIELD]: { rich_text: richText(firstValue(row, TASK_WORKSTREAM_FIELD)) },
    Status: { rich_text: richText(firstValue(row, "Status")) },
    Warning: { rich_text: richText(firstValue(row, "Warning")) },
    "Error Message": { rich_text: richText(firstValue(row, "Error Message")) },
    Active: { checkbox: parseBoolYesNo(firstValue(row, "Active")) },
    Placeholder: { checkbox: parseBoolYesNo(firstValue(row, "Placeholder")) },
    "Linked Fields": { checkbox: parseBoolYesNo(firstValue(row, "Linked Fields")) },
    [AUDIT_FIELD_SOURCE]: { rich_text: richText(audit.sourceFile) },
    [AUDIT_FIELD_VERSION]: { rich_text: richText(audit.version) },
    [AUDIT_FIELD_AT]: { date: { start: audit.importedAt } },
    [AUDIT_FIELD_ROW_HASH]: { rich_text: richText(rowHash) }
  };

  if (NOTES_OWNER === "msp") {
    properties.Notes = { rich_text: richText(firstValue(row, "Notes")) };
  }

  if (schemaContext.statusUpdateType === "status") {
    properties["Status Update"] = {
      status: { name: deriveStatusUpdate(percentComplete) }
    };
  }

  if (schemaContext.assigneeType === "people") {
    const emails = parseResourceList(firstValue(row, "Resource Email (Text1) (Text1)"))
      .map((email) => email.toLowerCase());
    const assigneeIds = [...new Set(
      emails
        .map((email) => schemaContext.userIdByEmail.get(email))
        .filter(Boolean)
    )];
    properties.Assignee = { people: assigneeIds.map((id) => ({ id })) };
  }

  return properties;
}

async function main() {
  assertEnv();
  ensureFolders();

  const sourceFile = pickPrimaryTaskCsv();
  const supplementalFile = findOptionalStagingFile(SUPPLEMENTAL_PLAN_NAME);
  const resourceTableFile = findOptionalStagingFile(RESOURCE_TABLE_NAME);
  const projectTeamFile = findOptionalStagingFile(PROJECT_TEAM_NAME);

  const fileName = path.basename(sourceFile);
  const { raw, rows } = parseCsvRows(sourceFile);
  const duplicateHeaders = listDuplicateHeaders(raw);
  const hash = crypto.createHash("sha1").update(raw).digest("hex").slice(0, 10);

  const supplementalByUid = parseSupplementalPlan(supplementalFile);
  const resourceTable = parseResourceTable(resourceTableFile);
  const projectTeam = parseProjectTeamRoster(projectTeamFile);
  mergeSupplementalIntoBaseRows(rows, supplementalByUid);
  const descriptionEnrichment = enrichWorkingTasksFromDescription(rows, projectTeam, resourceTable);

  if (!rows.length) {
    throw new Error("CSV contains no data rows.");
  }

  const stamp = new Date();
  const stampKey = stamp.toISOString().replace(/[:.]/g, "-");
  const importVersion = `${fileName}__${hash}__${stampKey}`;

  const audit = {
    sourceFile: fileName,
    version: importVersion,
    importedAt: stamp.toISOString().slice(0, 10)
  };

  console.log(`Using task CSV: ${sourceFile}`);
  if (supplementalFile) {
    console.log(`Using supplemental plan CSV: ${supplementalFile}`);
  }
  if (resourceTableFile) {
    console.log(`Using resource table CSV: ${resourceTableFile}`);
  }
  if (projectTeamFile) {
    console.log(`Using project team CSV: ${projectTeamFile}`);
  }
  console.log(
    `Description enrichment -> profiles: ${descriptionEnrichment.profiles}, working rows scanned: ${descriptionEnrichment.rowsConsidered}, ` +
      `Resource rows updated: ${descriptionEnrichment.resourceRowsUpdated}, BPO rows updated: ${descriptionEnrichment.bpoRowsUpdated}, Workstream rows updated: ${descriptionEnrichment.workstreamRowsUpdated}`
  );
  console.log(`Import version: ${importVersion}`);

  await ensureAuditProperties();

  const teamMembersDbId = await findTeamMembersDbId();

  let taskDb = await withRetry(
    () => notion.databases.retrieve({ database_id: TASKS_DB_ID }),
    "retrieve-task-db-schema"
  );

  await ensureTaskEnhancementProperties(taskDb.properties || {}, teamMembersDbId);

  taskDb = await withRetry(
    () => notion.databases.retrieve({ database_id: TASKS_DB_ID }),
    "retrieve-task-db-schema-after-ensure"
  );
  const taskDbProperties = taskDb.properties || {};

  const assigneeType = taskDbProperties.Assignee?.type || null;
  const statusUpdateType = taskDbProperties["Status Update"]?.type || null;

  let userIdByEmail = new Map();
  if (assigneeType === "people") {
    userIdByEmail = await fetchWorkspaceUsersByEmail();
    console.log(`Assignee user mapping loaded: ${userIdByEmail.size} users`);
  }

  const preflight = analyzeCsvPreflight(rows, duplicateHeaders);

  const existingPages = await fetchAllDatabasePages(TASKS_DB_ID, "task-pages");
  const pageByUniqueId = new Map();
  const pageIdByHierarchy = new Map();

  for (const page of existingPages) {
    const uniqueId = getNumberProperty(page, "Unique ID");
    if (uniqueId != null) {
      pageByUniqueId.set(uniqueId, page);
    }

    const hierarchy = getRichTextProperty(page, "Outline Number") || getRichTextProperty(page, "WBS");
    if (hierarchy) {
      pageIdByHierarchy.set(hierarchy, page.id);
    }
  }

  const projectPageId = await findAmetekProjectPageId();

  const teamSignals = collectTeamSignals(rows, resourceTable);
  const teamSync = await syncTeamMembers({
    teamMembersDbId,
    teamSignals,
    resourceTable,
    projectTeam,
    projectPageId
  });

  const existingRowHashByUniqueId = new Map();
  const existingRelationHashByUniqueId = new Map();
  for (const page of existingPages) {
    const uniqueId = getNumberProperty(page, "Unique ID");
    if (uniqueId == null) continue;
    existingRowHashByUniqueId.set(uniqueId, getRichTextProperty(page, AUDIT_FIELD_ROW_HASH));
    existingRelationHashByUniqueId.set(uniqueId, getRichTextProperty(page, AUDIT_FIELD_REL_HASH));
  }

  let created = 0;
  let updated = 0;
  let skipped = 0;
  let skippedUnchanged = 0;
  let relationSkippedUnchanged = 0;

  const rowMetaByUniqueId = new Map();
  const importRows = [];

  for (const row of rows) {
    const uniqueId = parseNumber(firstValue(row, "Unique ID"));
    if (uniqueId == null) {
      skipped += 1;
      continue;
    }
    importRows.push({ row, uniqueId });
  }

  await runWithConcurrency(
    importRows,
    IMPORT_CONCURRENCY,
    async ({ row, uniqueId }) => {
      const rowHash = buildRowHash(row);
      let properties = buildMspProperties(row, audit, rowHash, {
        assigneeType,
        statusUpdateType,
        userIdByEmail
      });
      properties["Unique ID"] = { number: uniqueId };
      properties = filterPropertiesBySchema(properties, taskDbProperties);

      const existing = pageByUniqueId.get(uniqueId);
      let pageId;

      if (existing) {
        const existingRowHash = existingRowHashByUniqueId.get(uniqueId) || "";
        if (IMPORT_MODE === "changed" && existingRowHash === rowHash) {
          skippedUnchanged += 1;
        } else {
          await withRetry(
            () => notion.pages.update({ page_id: existing.id, properties }),
            `upsert-update-${uniqueId}`
          );
          updated += 1;
        }
        pageId = existing.id;
      } else {
        const createdPage = await withRetry(
          () =>
            notion.pages.create({
              parent: { database_id: TASKS_DB_ID },
              properties
            }),
          `upsert-create-${uniqueId}`
        );
        pageId = createdPage.id;
        pageByUniqueId.set(uniqueId, { id: pageId });
        created += 1;
      }

      existingRowHashByUniqueId.set(uniqueId, rowHash);

      const hierarchyKey = getHierarchyKey(row);
      if (hierarchyKey) {
        pageIdByHierarchy.set(hierarchyKey, pageId);
      }

      rowMetaByUniqueId.set(uniqueId, {
        pageId,
        uniqueId,
        hierarchyKey,
        predecessors: parseIdList(firstValue(row, "Unique ID Predecessors")),
        successors: parseIdList(firstValue(row, "Unique ID Successors")),
        assigneeNames: splitNameList(firstValue(row, "Resource Names")),
        bpoNames: splitNameList(firstValue(row, "Business Validation Owner"))
      });
    },
    "Upsert phase"
  );

  let parentLinked = 0;
  let dependencyLinked = 0;
  let projectLinked = 0;
  let unresolvedParents = 0;
  let unresolvedDependencies = 0;

  const relationRows = Array.from(rowMetaByUniqueId.values());

  await runWithConcurrency(
    relationRows,
    IMPORT_CONCURRENCY,
    async (meta) => {
      const relationProps = {};
      let parentId = "";

      const parentKey = parentHierarchyKey(meta.hierarchyKey);
      if (parentKey) {
        const resolvedParentId = pageIdByHierarchy.get(parentKey);
        if (resolvedParentId) {
          parentId = resolvedParentId;
          relationProps["Parent Task"] = { relation: relationArray([resolvedParentId]) };
          parentLinked += 1;
        } else {
          unresolvedParents += 1;
        }
      } else {
        relationProps["Parent Task"] = { relation: [] };
      }

      const predecessorIds = meta.predecessors
        .map((id) => rowMetaByUniqueId.get(id)?.pageId || pageByUniqueId.get(id)?.id)
        .filter(Boolean);
      const distinctPredecessorIds = [...new Set(predecessorIds)];
      relationProps["Predecessor Tasks"] = {
        relation: relationArray(distinctPredecessorIds)
      };
      if (predecessorIds.length < meta.predecessors.length) {
        unresolvedDependencies += meta.predecessors.length - predecessorIds.length;
      }

      const successorIds = meta.successors
        .map((id) => rowMetaByUniqueId.get(id)?.pageId || pageByUniqueId.get(id)?.id)
        .filter(Boolean);
      const distinctSuccessorIds = [...new Set(successorIds)];
      relationProps["Successor Tasks"] = {
        relation: relationArray(distinctSuccessorIds)
      };
      if (successorIds.length < meta.successors.length) {
        unresolvedDependencies += meta.successors.length - successorIds.length;
      }

      dependencyLinked += predecessorIds.length + successorIds.length;

      if (projectPageId) {
        relationProps.Projects = { relation: relationArray([projectPageId]) };
        projectLinked += 1;
      }

      const assigneeTeamMemberIds = [...new Set(
        meta.assigneeNames
          .map((name) => teamSync.teamMemberIdByName.get(normalizeName(name)))
          .filter(Boolean)
      )];
      relationProps[TASK_ASSIGNEE_TEAM_REL] = {
        relation: relationArray(assigneeTeamMemberIds)
      };

      const bpoTeamMemberIds = [...new Set(
        meta.bpoNames
          .map((name) => teamSync.teamMemberIdByName.get(normalizeName(name)))
          .filter(Boolean)
      )];
      relationProps[TASK_BPO_TEAM_REL] = {
        relation: relationArray(bpoTeamMemberIds)
      };

      const relationHash = buildRelationHash({
        parentId,
        predecessorIds: distinctPredecessorIds,
        successorIds: distinctSuccessorIds,
        projectPageId,
        assigneeTeamMemberIds,
        bpoTeamMemberIds
      });
      relationProps[AUDIT_FIELD_REL_HASH] = { rich_text: richText(relationHash) };

      const safeRelationProps = filterPropertiesBySchema(relationProps, taskDbProperties);

      const existingRelationHash = existingRelationHashByUniqueId.get(meta.uniqueId) || "";
      if (IMPORT_MODE === "changed" && existingRelationHash === relationHash) {
        relationSkippedUnchanged += 1;
        return;
      }

      await withRetry(
        () => notion.pages.update({ page_id: meta.pageId, properties: safeRelationProps }),
        `relations-update-${meta.uniqueId}`
      );
      existingRelationHashByUniqueId.set(meta.uniqueId, relationHash);
    },
    "Relation phase"
  );

  const archiveName = `${stampKey}__${fileName}`;
  const archivePath = path.join(ARCHIVE_DIR, archiveName);
  fs.renameSync(sourceFile, archivePath);

  const report = {
    runtime: {
      ranAt: stamp.toISOString(),
      stampKey,
      mode: IMPORT_MODE,
      sourceFile,
      supplementalFile: supplementalFile || null,
      resourceTableFile: resourceTableFile || null,
      projectTeamFile: projectTeamFile || null,
      importVersion,
      notesOwner: NOTES_OWNER,
      importConcurrency: IMPORT_CONCURRENCY,
      requestDelayMs: REQUEST_DELAY_MS,
      apiTimeoutMs: API_TIMEOUT_MS
    },
    descriptionEnrichment,
    preflight,
    teamMembers: {
      databaseId: teamMembersDbId,
      syncedTotal: teamSync.total,
      created: teamSync.created,
      updated: teamSync.updated
    },
    importResult: {
      created,
      updated,
      skippedMissingUniqueId: skipped,
      skippedUnchangedRows: skippedUnchanged,
      parentLinksSet: parentLinked,
      dependencyLinksSet: dependencyLinked,
      projectLinksSet: projectLinked,
      skippedUnchangedRelations: relationSkippedUnchanged,
      unresolvedParentLinks: unresolvedParents,
      unresolvedDependencyRefs: unresolvedDependencies,
      archivedFile: archivePath
    }
  };
  writeImportReport(report);

  console.log("\nImport complete:");
  console.log(`  Mode: ${IMPORT_MODE}`);
  console.log(`  Created: ${created}`);
  console.log(`  Updated: ${updated}`);
  console.log(`  Skipped (missing Unique ID): ${skipped}`);
  console.log(`  Skipped unchanged (rows): ${skippedUnchanged}`);
  console.log(`  Team Members synced: ${teamSync.total} (created ${teamSync.created}, updated ${teamSync.updated})`);
  console.log(`  Parent links set: ${parentLinked}`);
  console.log(`  Dependency links set: ${dependencyLinked}`);
  console.log(`  Project links set: ${projectLinked}`);
  console.log(`  Skipped unchanged (relations): ${relationSkippedUnchanged}`);
  console.log(`  Unresolved parent links: ${unresolvedParents}`);
  console.log(`  Unresolved dependency refs: ${unresolvedDependencies}`);
  console.log(`  Archived file: ${archivePath}`);
  console.log(`  Audit version: ${importVersion}`);
  console.log(`  Report file: ${path.join(LOGS_DIR, "latest-import-report.json")}`);
}

main().catch((error) => {
  console.error("Import failed:", error?.body || error?.message || error);
  if (error?.stack) {
    console.error(error.stack);
  }
  process.exit(1);
});

/* eslint-disable no-console */
const fs = require("node:fs");
const path = require("node:path");
const { parse } = require("csv-parse/sync");

const ROOT = path.resolve(__dirname, "..");
const STAGING_DIR = path.join(ROOT, "imports", "staging");
const DOCS_DIR = path.join(ROOT, "docs");
const OUT_HTML = path.join(DOCS_DIR, "schedule-intelligence.html");
const OUT_JSON = path.join(DOCS_DIR, "schedule-intelligence-data.json");
const OUT_MERMAID = path.join(DOCS_DIR, "schedule-intelligence-critical.mmd");

const SUPPLEMENTAL_PLAN_NAME = "Project plan with resources and busines.csv";
const RESOURCE_TABLE_NAME = "Resource table.csv";
const PROJECT_TEAM_NAME = "project team.csv";
const TASK_WORKSTREAM_FIELD = "Workstream";

function ensureDir(dir) {
  fs.mkdirSync(dir, { recursive: true });
}

function canonical(text) {
  return String(text || "")
    .replace(/\s+/g, " ")
    .trim();
}

function normalizeName(value) {
  return canonical(value).toLowerCase();
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
    const safe = n.replace(/[.*+?^${}()|[\]\\]/g, "\\$&");
    return new RegExp(`\\b${safe}\\b`, "i").test(h);
  }
  return h.includes(n);
}

function firstValue(row, key) {
  const raw = row[key];
  if (Array.isArray(raw)) {
    for (const item of raw) {
      const text = canonical(item);
      if (text) return text;
    }
    return "";
  }
  return canonical(raw);
}

function parseNumber(value) {
  if (value == null) return null;
  const cleaned = String(value).replace(/,/g, "").replace(/%/g, "").trim();
  if (!cleaned) return null;
  const n = Number(cleaned);
  return Number.isFinite(n) ? n : null;
}

function parseBoolYesNo(value) {
  return /^(yes|true|1|y)$/i.test(String(value || "").trim());
}

function splitDelimited(value) {
  return String(value || "")
    .split(/[;,|/]/)
    .map((x) => canonical(x))
    .filter(Boolean);
}

const WORKSTREAM_ALIASES = new Map([
  ["basis", "Basis"],
  ["data", "Data"],
  ["development", "Development"],
  ["finance", "Finance"],
  ["functional", "Functional"],
  ["otc", "OTC"],
  ["pmo", "PMO"],
  ["ptm", "PTM"],
  ["test", "Testing"],
  ["testing", "Testing"],
  ["europe it lead", "Europe IT Lead"]
]);

function normalizeWorkstreamValue(value) {
  const firstSegment = splitDelimited(value)[0] || canonical(value);
  if (!firstSegment) return "";
  const text = canonical(firstSegment);
  const key = text.toLowerCase();
  if (WORKSTREAM_ALIASES.has(key)) return WORKSTREAM_ALIASES.get(key);
  if (/test/i.test(key)) return "Testing";
  if (/^\(unassigned\)$/i.test(text)) return "(Unassigned)";
  const normalized = text
    .split(/\s+/)
    .map((w) => {
      if (!w) return w;
      if (/^[A-Z0-9&/-]{2,5}$/.test(w)) return w.toUpperCase();
      return w.charAt(0).toUpperCase() + w.slice(1).toLowerCase();
    })
    .join(" ")
    .replace(/\bIt\b/g, "IT");
  return normalized;
}

function choosePrimaryWorkstream(values) {
  for (const v of values || []) {
    const ws = normalizeWorkstreamValue(v);
    if (ws) return ws;
  }
  return "";
}

function isMeaningfulValue(value) {
  const text = canonical(value);
  if (!text) return false;
  return !/could not be determined|tbd|unknown|unassigned/i.test(text);
}

function parseIdList(value) {
  const matches = String(value || "").match(/\d+/g);
  if (!matches) return [];
  return [...new Set(matches.map((m) => Number(m)).filter(Number.isFinite))];
}

function parseDateValue(value) {
  const v = canonical(value);
  if (!v || /^na$/i.test(v)) return null;
  const d = new Date(v);
  if (Number.isNaN(d.getTime())) return null;
  return d.toISOString().slice(0, 10);
}

function parseDurationDays(value) {
  const raw = canonical(value).replace(/\?/g, "");
  if (!raw || /^na$/i.test(raw)) return 0;
  const direct = Number(raw);
  if (Number.isFinite(direct)) return direct;

  const matches = [...raw.matchAll(/(-?\d+(?:\.\d+)?)\s*([a-z]+)/gi)];
  if (!matches.length) return 0;
  let totalDays = 0;
  for (const [, numText, unitRaw] of matches) {
    const num = Number(numText);
    const unit = unitRaw.toLowerCase();
    if (!Number.isFinite(num)) continue;
    if (unit.startsWith("w")) totalDays += num * 5;
    else if (unit.startsWith("d")) totalDays += num;
    else if (unit.startsWith("h")) totalDays += num / 8;
    else if (unit.startsWith("m")) totalDays += num / (8 * 60);
  }
  return totalDays;
}

function esc(value) {
  return String(value ?? "")
    .replace(/&/g, "&amp;")
    .replace(/</g, "&lt;")
    .replace(/>/g, "&gt;")
    .replace(/\"/g, "&quot;")
    .replace(/'/g, "&#39;");
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

function listStagingCsvFiles() {
  return fs.readdirSync(STAGING_DIR)
    .filter((name) => name.toLowerCase().endsWith(".csv"))
    .map((name) => ({
      name,
      full: path.join(STAGING_DIR, name),
      mtimeMs: fs.statSync(path.join(STAGING_DIR, name)).mtimeMs
    }))
    .sort((a, b) => b.mtimeMs - a.mtimeMs);
}

function detectTaskCsvScore(filePath) {
  try {
    const { rows } = parseCsvRows(filePath);
    if (!rows.length) return 0;
    const headers = Object.keys(rows[0]);
    const ok = headers.includes("Unique ID") && (headers.includes("Task Name") || headers.includes("Task description"));
    if (!ok) return 0;
    let score = 10;
    if (/msp task/i.test(filePath)) score += 10;
    if (/replan/i.test(filePath)) score += 5;
    return score;
  } catch {
    return 0;
  }
}

function pickPrimaryTaskCsv() {
  const scored = listStagingCsvFiles()
    .map((f) => ({ ...f, score: detectTaskCsvScore(f.full) }))
    .filter((f) => f.score > 0)
    .sort((a, b) => b.score - a.score || b.mtimeMs - a.mtimeMs);
  if (!scored.length) throw new Error(`No valid task CSV found in ${STAGING_DIR}`);
  return scored[0].full;
}

function findOptionalStagingFile(targetName) {
  const hit = listStagingCsvFiles().find((f) => f.name.toLowerCase() === targetName.toLowerCase());
  return hit?.full || null;
}

function parseMatrixCsv(filePath) {
  const raw = fs.readFileSync(filePath, "utf8");
  return parse(raw, {
    columns: false,
    bom: true,
    skip_empty_lines: false,
    trim: true,
    relax_column_count: true
  });
}

function parseSupplementalPlan(filePath) {
  const out = new Map();
  if (!filePath || !fs.existsSync(filePath)) return out;
  const matrix = parseMatrixCsv(filePath);
  const headerIndex = matrix.findIndex((row) => Array.isArray(row) && row.includes("Unique ID") && row.includes("Task description"));
  if (headerIndex < 0) return out;
  const header = matrix[headerIndex];
  const idx = {
    uniqueId: header.indexOf("Unique ID"),
    resources: header.indexOf("Resources"),
    businessValidation: header.indexOf("Business Validation"),
    workstream: header.indexOf("Workstream"),
    taskDescription: header.indexOf("Task description")
  };
  for (let i = headerIndex + 2; i < matrix.length; i += 1) {
    const row = matrix[i] || [];
    const uid = parseNumber(row[idx.uniqueId]);
    if (uid == null) continue;
    out.set(uid, {
      taskName: canonical(row[idx.taskDescription]),
      resourceNames: canonical(row[idx.resources]),
      businessValidationOwner: canonical(row[idx.businessValidation]),
      workstream: normalizeWorkstreamValue(row[idx.workstream])
    });
  }
  return out;
}

function parseResourceTable(filePath) {
  const result = { headers: [], rowsByName: new Map() };
  if (!filePath || !fs.existsSync(filePath)) return result;
  const matrix = parseMatrixCsv(filePath);
  const headerIndex = matrix.findIndex((row) => Array.isArray(row) && row.includes("Resource name") && row.includes("Abbreviation"));
  if (headerIndex < 0) return result;
  const header = matrix[headerIndex].map(canonical);
  result.headers = header.filter(Boolean);
  const idxByName = new Map();
  header.forEach((h, i) => idxByName.set(h, i));
  for (let i = headerIndex + 1; i < matrix.length; i += 1) {
    const row = matrix[i] || [];
    const name = canonical(row[idxByName.get("Resource name")]);
    if (!name) continue;
    const key = normalizeName(name);
    const existing = result.rowsByName.get(key) || { name, values: {}, workstreams: new Set() };
    for (const h of result.headers) {
      const idx = idxByName.get(h);
      const value = canonical(row[idx]);
      if (!value) continue;
      if (h === "Workstream") existing.workstreams.add(value);
      else if (!existing.values[h]) existing.values[h] = value;
    }
    result.rowsByName.set(key, existing);
  }
  return result;
}

function parseProjectTeamRoster(filePath) {
  const result = { headers: [], rowsByName: new Map() };
  if (!filePath || !fs.existsSync(filePath)) return result;
  const matrix = parseMatrixCsv(filePath);
  const headerIndex = matrix.findIndex((row) => {
    const cells = (row || []).map(canonical);
    return cells.includes("Name") && cells.some((c) => c.toLowerCase().includes("email")) && cells.some((c) => c.toLowerCase().includes("project position"));
  });
  if (headerIndex < 0) return result;
  const header = (matrix[headerIndex] || []).map(canonical);
  const idxByHeader = new Map();
  header.forEach((h, i) => { if (h) idxByHeader.set(h, i); });
  result.headers = [...idxByHeader.keys()];
  const nameIdx = idxByHeader.get("Name");
  if (nameIdx == null) return result;
  for (let i = headerIndex + 1; i < matrix.length; i += 1) {
    const row = matrix[i] || [];
    const name = canonical(row[nameIdx]);
    if (!name || name.toLowerCase().includes("if not assign")) continue;
    const key = normalizeName(name);
    const values = {};
    for (const h of result.headers) {
      const idx = idxByHeader.get(h);
      const value = canonical(row[idx]);
      if (value) values[h] = value;
    }
    result.rowsByName.set(key, { name, values });
  }
  return result;
}

function mergeSupplementalIntoBaseRows(baseRows, supplementalByUid) {
  for (const row of baseRows) {
    const uid = parseNumber(row.uid);
    if (uid == null) continue;
    const sup = supplementalByUid.get(uid);
    if (!sup) continue;
    if (sup.workstream && !isMeaningfulValue(row.workstream)) row.workstream = normalizeWorkstreamValue(sup.workstream);
    if (!row.name && sup.taskName) row.name = sup.taskName;
    if (!row.resourceNames && sup.resourceNames) row.resourceNames = sup.resourceNames;
    if (!row.businessValidationOwner && sup.businessValidationOwner) row.businessValidationOwner = sup.businessValidationOwner;
  }
}

function buildAssignmentProfiles(projectTeam, resourceTable) {
  const profiles = new Map();
  function ensureProfile(name) {
    const key = normalizeName(name);
    if (!key) return null;
    if (!profiles.has(key)) {
      profiles.set(key, {
        name: canonical(name),
        active: true,
        workstreams: new Set(),
        bpoOwners: new Set(),
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
    const status = canonical(row.values?.Status).toLowerCase();
    if (status && status !== "active") p.active = false;
    for (const term of splitDelimited(row.values?.["Project Team Assignment"])) {
      p.assignmentTerms.add(term); p.workstreams.add(term);
    }
    for (const term of splitDelimited(row.values?.["Project Position"])) p.positionTerms.add(term);
    if (/\bbpo\b|business process owner/i.test(`${row.values?.["Project Team Assignment"] || ""} ${row.values?.["Project Position"] || ""}`)) {
      p.isBpoLike = true;
    }
  }
  for (const row of resourceTable.rowsByName.values()) {
    const p = ensureProfile(row.name);
    if (!p) continue;
    for (const ws of row.workstreams || []) p.workstreams.add(ws);
    for (const owner of splitDelimited(row.values?.BPO)) {
      p.bpoOwners.add(owner);
      p.isBpoLike = true;
    }
  }
  return [...profiles.values()].filter((p) => p.active !== false);
}

function isWorkingTask(task) {
  return !task.summary && !task.milestone && !task.placeholder && task.active && task.percentComplete < 100;
}

function calculateSlipDays(task) {
  if (!task.finish || !task.baselineFinish) return 0;
  const slip = Math.round((new Date(task.finish) - new Date(task.baselineFinish)) / 86400000);
  return slip > 0 ? slip : 0;
}

function inferCycle(task) {
  const outline = String(task.outlineNumber || "");
  const name = String(task.name || "");
  if (outline.startsWith("1.2.1") || outline.startsWith("1.3.3.5") || outline.startsWith("1.1.1.4") || /\bITC[- ]?1\b/i.test(name)) return "ITC1";
  if (outline.startsWith("1.2.2") || outline.startsWith("1.3.3.6") || outline.startsWith("1.1.1.5") || /\bITC[- ]?2\b/i.test(name)) return "ITC2";
  if (outline.startsWith("1.2.3") || outline.startsWith("1.3.3.7") || outline.startsWith("1.1.1.6") || /\bITC[- ]?3\b/i.test(name)) return "ITC3";
  if (outline.startsWith("1.2.4") || outline.startsWith("1.3.3.8") || outline.startsWith("1.1.1.7") || /\bUAT\b/i.test(name)) return "UAT";
  return null;
}

function inferTrack(task) {
  const outline = String(task.outlineNumber || "");
  const name = String(task.name || "");
  if (outline.startsWith("1.1.1.") || /refresh/i.test(name)) return "Refresh";
  if (outline.startsWith("1.2.") || /data load|transactional|master data|simulate load|\bload\b/i.test(name)) return "Data Loads";
  if (outline.startsWith("1.3.3.") || /testing|test readiness|trr|gate review|defect/i.test(name)) return "Testing";
  return null;
}

function buildLaneAnchors(tasks) {
  const defs = [
    { id: "ITC1_REFRESH", cycle: "ITC1", track: "Refresh", label: "ITC1 Refresh", pattern: /QS5 ITC-1 Refresh/i, outlinePrefix: "1.1.1.4" },
    { id: "ITC2_REFRESH", cycle: "ITC2", track: "Refresh", label: "ITC2 Refresh", pattern: /QS4 ITC-2 Refresh/i, outlinePrefix: "1.1.1.5" },
    { id: "ITC3_REFRESH", cycle: "ITC3", track: "Refresh", label: "ITC3 Refresh", pattern: /QS5 ITC-3 Refresh/i, outlinePrefix: "1.1.1.6" },
    { id: "UAT_REFRESH", cycle: "UAT", track: "Refresh", label: "UAT Refresh", pattern: /QS4 UAT Refresh/i, outlinePrefix: "1.1.1.7" },
    { id: "ITC1_LOADS", cycle: "ITC1", track: "Data Loads", label: "ITC1 Data Loads", pattern: /ITC1 Data Load Tasks/i, outlinePrefix: "1.2.1" },
    { id: "ITC2_LOADS", cycle: "ITC2", track: "Data Loads", label: "ITC2 Data Loads", pattern: /ITC2 Data Load Tasks/i, outlinePrefix: "1.2.2" },
    { id: "ITC3_LOADS", cycle: "ITC3", track: "Data Loads", label: "ITC3 Data Loads", pattern: /ITC3 Data Load Tasks/i, outlinePrefix: "1.2.3" },
    { id: "UAT_LOADS", cycle: "UAT", track: "Data Loads", label: "UAT Data Loads", pattern: /UAT Data Load Tasks/i, outlinePrefix: "1.2.4" },
    { id: "ITC1_TESTING", cycle: "ITC1", track: "Testing", label: "ITC1 Testing", pattern: /^ITC-1$/i, outlinePrefix: "1.3.3.5" },
    { id: "ITC2_TESTING", cycle: "ITC2", track: "Testing", label: "ITC2 Testing", pattern: /^ITC-2$/i, outlinePrefix: "1.3.3.6" },
    { id: "ITC3_TESTING", cycle: "ITC3", track: "Testing", label: "ITC3 Testing", pattern: /^ITC-3$/i, outlinePrefix: "1.3.3.7" },
    { id: "UAT_TESTING", cycle: "UAT", track: "Testing", label: "UAT Testing", pattern: /^UAT \(DR\)$/i, outlinePrefix: "1.3.3.8" }
  ];

  return defs.map((def) => {
    const anchor = tasks.find((t) => t.summary && def.pattern.test(t.name))
      || tasks.find((t) => t.summary && String(t.outlineNumber || "").startsWith(def.outlinePrefix));
    if (!anchor) return null;
    return {
      ...def,
      uid: anchor.uid,
      name: anchor.name,
      outlineNumber: anchor.outlineNumber,
      outlineLevel: anchor.outlineLevel,
      start: anchor.start,
      finish: anchor.finish
    };
  }).filter(Boolean);
}

function rankMilestoneRisk(a, b) {
  const aPast = a.daysToFinish != null && a.daysToFinish < 0 ? 1 : 0;
  const bPast = b.daysToFinish != null && b.daysToFinish < 0 ? 1 : 0;
  return (b.critical ? 1 : 0) - (a.critical ? 1 : 0)
    || bPast - aPast
    || (b.slipDays || 0) - (a.slipDays || 0)
    || (a.daysToFinish ?? 99999) - (b.daysToFinish ?? 99999)
    || String(a.name).localeCompare(String(b.name));
}

function summarizeTaskCard(task) {
  if (!task) return null;
  return {
    uid: task.uid,
    name: task.name,
    outlineNumber: task.outlineNumber,
    finish: task.finish,
    baselineFinish: task.baselineFinish,
    slipDays: calculateSlipDays(task),
    cycle: inferCycle(task),
    track: inferTrack(task),
    workstream: task.workstream,
    resourceNames: task.resourceNames,
    critical: !!(task.cpm && task.cpm.critical),
    driverNames: (task.drivers || []).map((d) => d.name).slice(0, 3)
  };
}

function deriveMilestoneParent(task) {
  const text20 = canonical(task?.text20);
  if (text20) return text20;
  return inferCycle(task) || task?.workstream || "Program";
}

/**
 * Build critical chain narrative: extract sequence of critical tasks,
 * identify which are at-risk (IMMEDIATE/SOON), compute cascade impact.
 */
function buildCriticalChainNarrative(tasks, byUid) {
  const today = new Date(); today.setHours(0, 0, 0, 0);
  
  // Get all critical path tasks, sorted by early start
  const criticalTasks = tasks
    .filter((t) => t.cpm?.critical && !t.summary && t.active)
    .sort((a, b) => (a.cpm.earlyStart - b.cpm.earlyStart) || a.uid - b.uid);

  if (criticalTasks.length === 0) return null;

  // Find start (first critical task)
  const startTask = criticalTasks[0];
  const finishTask = criticalTasks[criticalTasks.length - 1];

  // For each critical task, check if at-risk
  const chainItems = criticalTasks.map((task) => {
    const slipDays = calculateSlipDays(task);
    const finish = task.finish ? new Date(task.finish) : null;
    const daysToFinish = finish && finish >= today
      ? Math.round((finish - today) / 86400000)
      : -1;

    // Classify risk on critical path
    let riskStatus = "ON_TRACK";
    if (slipDays >= 30) riskStatus = "CRITICAL_SLIP";
    else if (slipDays >= 7) riskStatus = "SLIPPING";
    else if (slipDays > 0) riskStatus = "MINOR_SLIP";
    else if (daysToFinish < 0) riskStatus = "OVERDUE";

    // Count immediate downstream successors
    const directDownstream = (task.successors || [])
      .map((uid) => byUid.get(uid))
      .filter(Boolean)
      .filter((t) => !t.summary && t.active)
      .length;

    return {
      uid: task.uid,
      name: task.name,
      outlineNumber: task.outlineNumber,
      start: task.start,
      finish: task.finish,
      slipDays,
      daysToFinish,
      riskStatus,
      workstream: task.workstream,
      resourceNames: task.resourceNames,
      totalSlack: task.cpm?.totalSlack,
      successorCount: directDownstream
    };
  });

  // Cascade impact: if first critical task slips 5 days, how many tasks affected?
  const cascadeDownstream = (task, slipDays) => {
    const visited = new Set();
    const queue = [task.uid];
    let count = 0;
    const levels = { direct: 0, secondary: 0 };

    while (queue.length > 0 && count < 200) {
      const uid = queue.shift();
      if (visited.has(uid)) continue;
      visited.add(uid);

      const successor = byUid.get(uid);
      if (!successor || successor.summary || !successor.active) continue;

      count += 1;
      if (count <= (successor.successors || []).length) levels.direct += 1;
      else levels.secondary += 1;

      (successor.successors || []).forEach((sUid) => {
        if (!visited.has(sUid)) queue.push(sUid);
      });
    }
    return { count, ...levels };
  };

  // Analyze at-risk critical tasks
  const atRiskItems = chainItems.filter((c) =>
    c.riskStatus !== "ON_TRACK" && c.riskStatus !== "MINOR_SLIP"
  );

  const projectDuration = finishTask.cpm?.earlyFinish - startTask.cpm?.earlyStart;
  const projectEndDate = finishTask.finish;

  return {
    startTask: chainItems[0],
    finishTask: chainItems[chainItems.length - 1],
    chain: chainItems,
    atRiskCount: atRiskItems.length,
    projectDuration,
    projectEndDate,
    cascadeImpactIfFirstAtRiskSlips: atRiskItems.length > 0
      ? cascadeDownstream(atRiskItems[0], atRiskItems[0].slipDays)
      : null
  };
}

/**
 * Build triage guide: rank IMMEDIATE tasks by criticality.
 * Tier 1: On critical path (critical=true)
 * Tier 2: Blocking critical path (has critical successor)
 * Tier 3: Other IMMEDIATE
 */
function buildTriageGuide(tasks, immediate, byUid) {
  if (!immediate || immediate.length === 0) return null;

  const criticalUids = new Set(
    tasks
      .filter((t) => t.cpm?.critical && !t.summary && t.active)
      .map((t) => t.uid)
  );

  // Mark each task with triage tier
  const tiered = immediate.map((task) => {
    let triageTier = "TIER3_OTHER";

    // Tier 1: On critical path
    if (criticalUids.has(task.uid)) {
      triageTier = "TIER1_CRITICAL";
    }
    // Tier 2: Blocking critical path
    else if (task.successors && task.successors.length > 0) {
      const blocksCritical = task.successors.some((uid) => criticalUids.has(uid));
      if (blocksCritical) triageTier = "TIER2_BLOCKS_CRITICAL";
    }

    return { ...task, triageTier };
  });

  // Sort: Tier 1 first (by slip days desc), then Tier 2, then Tier 3
  const sorted = tiered.sort((a, b) => {
    const tierOrder = { TIER1_CRITICAL: 0, TIER2_BLOCKS_CRITICAL: 1, TIER3_OTHER: 2 };
    const tierCmp = tierOrder[a.triageTier] - tierOrder[b.triageTier];
    if (tierCmp !== 0) return tierCmp;
    return (b.slipDays || 0) - (a.slipDays || 0);
  });

  return {
    tier1Count: sorted.filter((t) => t.triageTier === "TIER1_CRITICAL").length,
    tier2Count: sorted.filter((t) => t.triageTier === "TIER2_BLOCKS_CRITICAL").length,
    tier3Count: sorted.filter((t) => t.triageTier === "TIER3_OTHER").length,
    topTier1: sorted.filter((t) => t.triageTier === "TIER1_CRITICAL").slice(0, 3),
    topTier2: sorted.filter((t) => t.triageTier === "TIER2_BLOCKS_CRITICAL").slice(0, 3),
    guide: sorted.slice(0, 12)
  };
}

function buildDecisionMetrics(tasks, prioritized) {
  const today = new Date();
  today.setHours(0, 0, 0, 0);

  const openMilestones = tasks
    .filter((t) => t.milestone && t.active && t.percentComplete < 100 && t.finish)
    .map((t) => {
      const finish = new Date(t.finish);
      return {
        ...t,
        cycle: inferCycle(t),
        track: inferTrack(t),
        slipDays: calculateSlipDays(t),
        daysToFinish: Number.isNaN(finish.getTime()) ? null : Math.round((finish - today) / 86400000),
        critical: !!(t.cpm && t.cpm.critical)
      };
    });

  const nextMilestone = [...openMilestones]
    .filter((t) => t.daysToFinish != null && t.daysToFinish >= 0)
    .sort((a, b) => a.daysToFinish - b.daysToFinish || rankMilestoneRisk(a, b))[0] || null;

  // Smart scroll date: focus on next at-risk milestone
  let smartScrollDate = null;
  const atRiskMilestones = [...openMilestones].filter((t) => t.critical || t.slipDays > 0 || t.daysToFinish < 14).sort(rankMilestoneRisk);
  if (atRiskMilestones.length) {
    smartScrollDate = atRiskMilestones[0].start || atRiskMilestones[0].finish;
  } else if (nextMilestone) {
    smartScrollDate = nextMilestone.start || nextMilestone.finish;
  }

  const topDataLoadRisk = [...openMilestones]
    .filter((t) => t.track === "Data Loads")
    .sort(rankMilestoneRisk)[0] || null;

  const topTestingRisk = [...openMilestones]
    .filter((t) => t.track === "Testing")
    .sort(rankMilestoneRisk)[0] || null;

  const topRefreshRisk = [...openMilestones]
    .filter((t) => t.track === "Refresh")
    .sort(rankMilestoneRisk)[0] || null;

  const byUid = new Map(tasks.map((t) => [t.uid, t]));
  const milestoneCommandTable = [...openMilestones]
    .sort(rankMilestoneRisk)
    .slice(0, 24)
    .map((t) => {
      const succCount = (byUid.get(t.uid)?.successors || []).length;
      const status = t.critical
        ? "Critical"
        : t.slipDays >= 30
          ? "Off-track"
          : t.slipDays > 0
            ? "Slipping"
            : (t.daysToFinish != null && t.daysToFinish < 0)
              ? "Overdue"
              : "On-track";
      return {
        uid: t.uid,
        parentMilestone: deriveMilestoneParent(t),
        cycle: t.cycle,
        track: t.track,
        name: t.name,
        finish: t.finish,
        baselineFinish: t.baselineFinish,
        daysToFinish: t.daysToFinish,
        slipDays: t.slipDays,
        critical: t.critical,
        status,
        impact: succCount
      };
    });

  const focusExamples = {
    refresh: [...openMilestones].filter((t) => t.cycle === "ITC2" && t.track === "Refresh").sort(rankMilestoneRisk)[0] || null,
    loads: [...openMilestones].filter((t) => t.cycle === "ITC2" && t.track === "Data Loads").sort(rankMilestoneRisk)[0] || null,
    testing: prioritized.filter((t) => inferCycle(t) === "ITC2" && inferTrack(t) === "Testing").sort((a, b) => b.slipDays - a.slipDays)[0] || null
  };

  const immediateExample = prioritized.find((t) => inferCycle(t) === "ITC2" && inferTrack(t) === "Data Loads")
    || prioritized.find((t) => inferCycle(t) === "ITC2")
    || prioritized[0]
    || null;

  return {
    nextMilestone: summarizeTaskCard(nextMilestone),
    smartScrollDate,
    topDataLoadRisk: summarizeTaskCard(topDataLoadRisk),
    topTestingRisk: summarizeTaskCard(topTestingRisk),
    topRefreshRisk: summarizeTaskCard(topRefreshRisk),
    immediateExample: immediateExample ? {
      uid: immediateExample.uid,
      name: immediateExample.name,
      outlineNumber: immediateExample.outlineNumber,
      finish: immediateExample.finish,
      slipDays: immediateExample.slipDays,
      tier: immediateExample.tier,
      resourceNames: immediateExample.resourceNames,
      workstream: immediateExample.workstream,
      why: immediateExample.why,
      impact: immediateExample.impact,
      resolution: immediateExample.resolution
    } : null,
    focusExamples: {
      refresh: summarizeTaskCard(focusExamples.refresh),
      loads: summarizeTaskCard(focusExamples.loads),
      testing: focusExamples.testing ? {
        uid: focusExamples.testing.uid,
        name: focusExamples.testing.name,
        outlineNumber: focusExamples.testing.outlineNumber,
        finish: focusExamples.testing.finish,
        slipDays: focusExamples.testing.slipDays,
        tier: focusExamples.testing.tier,
        resourceNames: focusExamples.testing.resourceNames,
        why: focusExamples.testing.why
      } : null
    },
    atRiskMilestones: openMilestones.filter((t) => t.slipDays > 0 || (t.daysToFinish != null && t.daysToFinish <= 30)).sort(rankMilestoneRisk).slice(0, 8).map(summarizeTaskCard),
    criticalChainNarrative: buildCriticalChainNarrative(tasks, new Map(tasks.map((t) => [t.uid, t]))),
    triageGuide: buildTriageGuide(tasks, prioritized.filter((t) => t.tier === "IMMEDIATE"), new Map(tasks.map((t) => [t.uid, t]))),
    milestoneCommandTable
  };
}

function pickMatches(taskText, profiles) {
  const scored = [];
  for (const p of profiles) {
    let score = 0;
    let matchedName = false;
    let matchedAssignment = false;
    let matchedWorkstream = false;
    if (containsPhrase(taskText, p.name)) {
      score += 12;
      matchedName = true;
    }
    for (const np of p.name.split(/\s+/).filter((x) => x.length >= 4)) {
      if (containsPhrase(taskText, np)) {
        score += 4;
        matchedName = true;
      }
    }
    for (const term of p.workstreams) {
      if (containsPhrase(taskText, term)) {
        score += 6;
        matchedWorkstream = true;
      }
    }
    for (const term of p.assignmentTerms) {
      if (containsPhrase(taskText, term)) {
        score += 5;
        matchedAssignment = true;
      }
    }
    for (const term of p.positionTerms) {
      if (containsPhrase(taskText, term)) {
        score += 2;
        matchedAssignment = true;
      }
    }
    if (score > 0) scored.push({ profile: p, score, matchedName, matchedAssignment, matchedWorkstream });
  }
  scored.sort((a, b) => b.score - a.score || a.profile.name.localeCompare(b.profile.name));
  return scored;
}

function enrichWorkingTasks(tasks, projectTeam, resourceTable) {
  const profiles = buildAssignmentProfiles(projectTeam, resourceTable);
  let enrichedResource = 0, enrichedBpo = 0, enrichedWorkstream = 0;
  for (const task of tasks) {
    if (!isWorkingTask(task)) continue;
    const matches = pickMatches(`${task.name} ${task.summaryName || ""}`, profiles).slice(0, 6);
    if (!matches.length) continue;
    const resources = [...new Set(matches
      .filter((m) => !m.profile.isBpoLike && (m.matchedName || m.matchedAssignment || m.score >= 10))
      .map((m) => m.profile.name))];
    const workstreams = [...new Set(matches
      .filter((m) => m.matchedWorkstream || m.matchedAssignment)
      .flatMap((m) => [...m.profile.workstreams]))];
    const bpos = [...new Set(matches
      .filter((m) => m.matchedWorkstream || m.matchedAssignment)
      .flatMap((m) => [...m.profile.bpoOwners]))];

    if (!isMeaningfulValue(task.resourceNames) && resources.length) {
      task.resourceNames = resources.slice(0, 3).join("; ");
      enrichedResource += 1;
    }
    if (!isMeaningfulValue(task.businessValidationOwner) && bpos.length) {
      task.businessValidationOwner = bpos.slice(0, 2).join("; ");
      enrichedBpo += 1;
    }
    if (!isMeaningfulValue(task.workstream) && workstreams.length) {
      task.workstream = choosePrimaryWorkstream(workstreams);
      enrichedWorkstream += 1;
    }
  }
  return { profiles: profiles.length, resourceRowsUpdated: enrichedResource, bpoRowsUpdated: enrichedBpo, workstreamRowsUpdated: enrichedWorkstream };
}

function normalizeTaskWorkstreams(tasks) {
  for (const task of tasks) {
    task.workstream = normalizeWorkstreamValue(task.workstream);
  }
}

function loadTasks(taskCsvPath) {
  const { rows } = parseCsvRows(taskCsvPath);
  return rows.map((row) => ({
    uid: parseNumber(firstValue(row, "Unique ID")),
    wbs: firstValue(row, "WBS"),
    outlineNumber: firstValue(row, "Outline Number") || firstValue(row, "WBS"),
    outlineLevel: parseNumber(firstValue(row, "Outline Level")) || 0,
    name: firstValue(row, "Task Name") || firstValue(row, "Task description"),
    summaryName: firstValue(row, "Task Summary Name"),
    text20: firstValue(row, "Text20") || firstValue(row, "Text 20"),
    start: parseDateValue(firstValue(row, "Start")),
    finish: parseDateValue(firstValue(row, "Finish")),
    baselineStart: parseDateValue(firstValue(row, "Baseline Start")) || parseDateValue(firstValue(row, "Baseline10 Start")),
    baselineFinish: parseDateValue(firstValue(row, "Baseline Finish")) || parseDateValue(firstValue(row, "Baseline10 Finish")),
    percentComplete: parseNumber(firstValue(row, "% Complete")) || 0,
    durationDays: parseDurationDays(firstValue(row, "Duration")),
    predecessors: parseIdList(firstValue(row, "Unique ID Predecessors")),
    successors: parseIdList(firstValue(row, "Unique ID Successors")),
    constraintType: firstValue(row, "Constraint Type"),
    status: firstValue(row, "Status"),
    totalSlackRaw: firstValue(row, "Total Slack"),
    freeSlackRaw: firstValue(row, "Free Slack"),
    resourceNames: firstValue(row, "Resource Names"),
    businessValidationOwner: firstValue(row, "Business Validation Owner"),
    workstream: normalizeWorkstreamValue(firstValue(row, TASK_WORKSTREAM_FIELD)),
    summary: parseBoolYesNo(firstValue(row, "Summary")),
    milestone: parseBoolYesNo(firstValue(row, "Milestone")),
    criticalFlag: parseBoolYesNo(firstValue(row, "Critical")),
    active: firstValue(row, "Active") ? parseBoolYesNo(firstValue(row, "Active")) : true,
    placeholder: parseBoolYesNo(firstValue(row, "Placeholder"))
  })).filter((t) => t.uid != null && t.name);
}

function buildCpm(tasks) {
  const nodes = tasks.filter((t) => !t.summary && !t.placeholder && t.active);
  const byUid = new Map(nodes.map((t) => [t.uid, t]));
  const indegree = new Map();
  const outgoing = new Map();
  for (const t of nodes) {
    indegree.set(t.uid, 0);
    outgoing.set(t.uid, []);
  }
  for (const t of nodes) {
    for (const pred of t.predecessors) {
      if (!byUid.has(pred)) continue;
      indegree.set(t.uid, (indegree.get(t.uid) || 0) + 1);
      outgoing.get(pred).push(t.uid);
    }
  }
  const queue = nodes.filter((t) => (indegree.get(t.uid) || 0) === 0).map((t) => t.uid);
  const topo = [];
  while (queue.length) {
    const uid = queue.shift();
    topo.push(uid);
    for (const next of outgoing.get(uid) || []) {
      indegree.set(next, indegree.get(next) - 1);
      if (indegree.get(next) === 0) queue.push(next);
    }
  }
  const cyclic = topo.length !== nodes.length;
  if (cyclic) {
    const missing = nodes.map((t) => t.uid).filter((uid) => !topo.includes(uid));
    topo.push(...missing);
  }

  const earlyStart = new Map();
  const earlyFinish = new Map();
  for (const uid of topo) {
    const t = byUid.get(uid);
    const preds = t.predecessors.filter((p) => byUid.has(p));
    const es = preds.length ? Math.max(...preds.map((p) => earlyFinish.get(p) || 0)) : 0;
    const ef = es + Math.max(t.durationDays || 0, t.milestone ? 0 : 0);
    earlyStart.set(uid, es);
    earlyFinish.set(uid, ef);
  }
  const projectDuration = topo.length ? Math.max(...topo.map((uid) => earlyFinish.get(uid) || 0)) : 0;
  const lateFinish = new Map();
  const lateStart = new Map();
  for (const uid of [...topo].reverse()) {
    const t = byUid.get(uid);
    const succs = (outgoing.get(uid) || []).filter((s) => byUid.has(s));
    const lf = succs.length ? Math.min(...succs.map((s) => lateStart.get(s) ?? projectDuration)) : projectDuration;
    const ls = lf - Math.max(t.durationDays || 0, t.milestone ? 0 : 0);
    lateFinish.set(uid, lf);
    lateStart.set(uid, ls);
  }
  const critical = [];
  for (const uid of topo) {
    const t = byUid.get(uid);
    const slack = (lateStart.get(uid) ?? 0) - (earlyStart.get(uid) ?? 0);
    t.cpm = {
      earlyStart: +(earlyStart.get(uid) || 0).toFixed(2),
      earlyFinish: +(earlyFinish.get(uid) || 0).toFixed(2),
      lateStart: +(lateStart.get(uid) || 0).toFixed(2),
      lateFinish: +(lateFinish.get(uid) || 0).toFixed(2),
      totalSlack: +slack.toFixed(2),
      critical: Math.abs(slack) < 0.01
    };
    if (t.cpm.critical) critical.push(t);
  }

  for (const t of nodes) {
    const preds = t.predecessors.filter((p) => byUid.has(p));
    const es = earlyStart.get(t.uid) || 0;
    const drivers = preds
      .map((p) => ({ uid: p, ef: earlyFinish.get(p) || 0, name: byUid.get(p)?.name || String(p) }))
      .filter((p) => Math.abs(p.ef - es) < 0.01)
      .sort((a, b) => a.name.localeCompare(b.name));
    t.drivers = drivers;
  }

  return { projectDuration: +projectDuration.toFixed(2), cyclic, taskCount: nodes.length, criticalCount: critical.length };
}

/**
 * Classify priority tier based on CPM output + scheduled start awareness.
 * Returns "IMMEDIATE" | "SOON" | "WATCH" | null (null = not a current concern).
 *
 * Key principle: a task starting > 4 weeks from today with no critical-path
 * flag is NOT urgent — even if it's slipped vs baseline. We care about what
 * needs a decision NOW, not every future risk.
 */
function classifyPriority(task) {
  const today = new Date(); today.setHours(0, 0, 0, 0);
  const start  = task.start  ? new Date(task.start)  : null;
  const finish = task.finish ? new Date(task.finish) : null;
  const baselineFinish = task.baselineFinish ? new Date(task.baselineFinish) : null;

  const slipDays = (finish && baselineFinish && finish > baselineFinish)
    ? Math.round((finish - baselineFinish) / 86400000) : 0;
  const daysToStart  = start  ? Math.round((start  - today) / 86400000) : null;
  const daysTilFinish = finish ? Math.round((finish - today) / 86400000) : null;
  const pastDue = finish && finish < today;
  const pastDueBaseline = baselineFinish && baselineFinish < today && task.percentComplete < 100;

  // ALWAYS IMMEDIATE if on CPM critical path (zero slack) — regardless of start date
  if (task.cpm?.critical) return "IMMEDIATE";

  // Near-critical (CPM slack ≤ 5 days) — flag regardless of start date
  if (task.cpm?.totalSlack != null && task.cpm.totalSlack <= 5) return "WATCH";

  const hasStarted     = daysToStart !== null && daysToStart <= 0;   // in progress
  const startsIn14     = daysToStart !== null && daysToStart <= 14;  // starts within 2 weeks
  const startsIn28     = daysToStart !== null && daysToStart <= 28;  // starts within 4 weeks

  // Task is active or starts within 2 weeks — full urgency rules apply
  if (hasStarted || startsIn14) {
    if (pastDue || pastDueBaseline) return "IMMEDIATE";
    if (slipDays >= 30) return "IMMEDIATE";
    if (slipDays >= 7 || (daysTilFinish != null && daysTilFinish <= 14)) return "SOON";
    if (task.cpm?.totalSlack != null && task.cpm.totalSlack <= 14) return "WATCH";
    return "WATCH"; // actively running, keep an eye on it
  }

  // Task starts in weeks 3–4 — moderate urgency
  if (startsIn28) {
    if (pastDueBaseline || slipDays >= 60) return "IMMEDIATE";
    if (slipDays >= 21 || (daysTilFinish != null && daysTilFinish <= 14)) return "SOON";
    if (slipDays > 0) return "WATCH";
    return null; // starts in ~3-4 weeks, no slip, fine
  }

  // Task starts > 4 weeks from today — only flag if seriously off-track
  if (slipDays >= 90) return "WATCH"; // major future slip worth watching
  return null; // future task, not a current concern
}

function buildTaskNarrative(task, byUid) {
  const slipDays = calculateSlipDays(task);
  const today = new Date(); today.setHours(0, 0, 0, 0);
  const finish = task.finish ? new Date(task.finish) : null;
  const pastDue = finish && finish < today;

  let whatIsWrong = "";
  if (task.cpm?.critical) whatIsWrong = "On critical path — any delay extends project end date";
  else if (pastDue) whatIsWrong = `Finish date ${task.finish} has passed (task not complete)`;
  else if (slipDays >= 90) whatIsWrong = `Schedule slipped ${slipDays} days behind baseline`;
  else if (slipDays >= 30) whatIsWrong = `Schedule slipping — ${slipDays} days behind baseline`;
  else if (slipDays > 0) whatIsWrong = `Minor slip — ${slipDays} days behind baseline`;
  else whatIsWrong = "Within 14-day window or low slack — needs monitoring";

  let why = "";
  if (task.drivers && task.drivers.length > 0) {
    why = "Driven by: " + task.drivers.slice(0, 3).map((d) => d.name).join("; ");
  } else if (task.predecessors && task.predecessors.length > 0) {
    const predNames = task.predecessors.slice(0, 3).map((p) => {
      const pred = byUid ? byUid.get(p) : null;
      return pred ? pred.name : `Task ${p}`;
    });
    why = "Predecessor dependency: " + predNames.join("; ");
  } else {
    why = "No predecessor chain — may be resource or scope driven";
  }

  const succCount = (task.successors || []).length;
  let impact = "";
  if (succCount === 0) impact = "No direct successors — isolated risk";
  else if (succCount <= 3) impact = `Blocks ${succCount} downstream task${succCount > 1 ? "s" : ""}`;
  else impact = `Blocks ${succCount} downstream tasks — broad schedule impact`;

  let resolution = "";
  if (task.cpm?.critical) resolution = "Escalate immediately; compress schedule or add resources";
  else if (slipDays >= 60) resolution = "Requires recovery plan; consider fast-tracking with successor tasks";
  else if (slipDays >= 14) resolution = "Engage task owner for revised plan and commitment date";
  else resolution = "Monitor daily; verify task is actively progressing";

  return { whatIsWrong, why, impact, resolution, slipDays };
}

function computeScheduleStats(tasks, cpmSummary) {
  const working = tasks.filter(isWorkingTask);
  const byUid = new Map(tasks.map((t) => [t.uid, t]));

  const slipped = working.filter((t) => {
    if (!t.finish || !t.baselineFinish) return false;
    return new Date(t.finish) > new Date(t.baselineFinish);
  }).map((t) => ({
    ...t,
    slipDays: Math.round((new Date(t.finish) - new Date(t.baselineFinish)) / 86400000)
  })).sort((a, b) => b.slipDays - a.slipDays);

  const today = new Date(); today.setHours(0, 0, 0, 0);
  const future14 = new Date(today); future14.setDate(future14.getDate() + 14);

  const due14 = working.filter((t) => {
    if (!t.finish) return false;
    const d = new Date(t.finish);
    return d >= today && d <= future14;
  }).sort((a, b) => String(a.finish).localeCompare(String(b.finish)));

  // Priority tiers — only working tasks; null tier = not a current concern
  const prioritized = working.map((t) => {
    const tier = classifyPriority(t);
    if (!tier) return null;
    const narrative = buildTaskNarrative(t, byUid);
    return { ...t, tier, ...narrative };
  }).filter(Boolean);

  const immediate = prioritized.filter((t) => t.tier === "IMMEDIATE").sort((a, b) => b.slipDays - a.slipDays || String(a.finish).localeCompare(String(b.finish)));
  const soon = prioritized.filter((t) => t.tier === "SOON").sort((a, b) => String(a.finish).localeCompare(String(b.finish)));
  const watch = prioritized.filter((t) => t.tier === "WATCH").sort((a, b) => (a.cpm?.totalSlack ?? 999) - (b.cpm?.totalSlack ?? 999));

  // Workstream health — only working tasks
  const wsHealth = new Map();
  for (const t of prioritized) {
    const ws = t.workstream || "(Unassigned)";
    if (!wsHealth.has(ws)) wsHealth.set(ws, { workstream: ws, total: 0, immediate: 0, soon: 0, watch: 0 });
    const rec = wsHealth.get(ws);
    rec.total += 1;
    if (t.tier === "IMMEDIATE") rec.immediate += 1;
    else if (t.tier === "SOON") rec.soon += 1;
    else rec.watch += 1;
  }
  const workstreamHealth = [...wsHealth.values()].sort((a, b) => b.immediate - a.immediate || b.total - a.total);
  const decisionMetrics = buildDecisionMetrics(tasks, prioritized);
  const criticalChain = (tasks || [])
    .filter((t) => t.cpm?.critical && !t.summary && t.active)
    .sort((a, b) => (a.cpm.earlyStart - b.cpm.earlyStart) || a.uid - b.uid);
  const cpmBreakdown = criticalChain.map((task, idx) => {
    const slipDays = calculateSlipDays(task);
    const drivers = (task.drivers || []).map((d) => d.name);
    const successorNames = (task.successors || []).map((uid) => byUid.get(uid)?.name).filter(Boolean);
    const narrative = buildTaskNarrative(task, byUid);
    return {
      rank: idx + 1,
      uid: task.uid,
      name: task.name,
      outlineNumber: task.outlineNumber,
      workstream: task.workstream,
      resourceNames: task.resourceNames,
      start: task.start,
      finish: task.finish,
      baselineFinish: task.baselineFinish,
      earlyStart: task.cpm?.earlyStart,
      earlyFinish: task.cpm?.earlyFinish,
      totalSlack: task.cpm?.totalSlack,
      slipDays,
      drivenBy: drivers,
      blocks: successorNames,
      what: narrative.whatIsWrong || [],
      why: narrative.why,
      impact: narrative.impact,
      action: narrative.resolution
    };
  });

  return {
    totalTasks: tasks.length,
    workingTasks: working.length,
    criticalTasks: working.filter((t) => t.cpm?.critical).length,
    nearCriticalTasks: working.filter((t) => !t.cpm?.critical && t.cpm?.totalSlack != null && t.cpm.totalSlack <= 10).length,
    slippedOpen: slipped.length,
    due14: due14.length,
    immediateCount: immediate.length,
    soonCount: soon.length,
    watchCount: watch.length,
    topWorkstreams: workstreamHealth.slice(0, 10).map(({ workstream, total }) => ({ workstream, count: total })),
    workstreamHealth,
    topSlipped: slipped.slice(0, 25),
    dueSoon: due14.slice(0, 25),
    immediate: immediate.slice(0, 100),
    soon: soon.slice(0, 100),
    watch: watch.slice(0, 100),
    cpmBreakdown,
    decisionMetrics,
    cpmSummary
  };
}

function buildMermaid(tasks) {
  const critical = tasks.filter((t) => t.cpm?.critical).sort((a, b) => (a.cpm.earlyStart - b.cpm.earlyStart) || a.uid - b.uid).slice(0, 40);
  const criticalIds = new Set(critical.map((t) => t.uid));
  const lines = ["flowchart LR"];
  for (const t of critical) {
    lines.push(`  T${t.uid}[\"${String(t.uid)} · ${String(t.name).replace(/\"/g, "'").slice(0, 42)}\"]`);
  }
  for (const t of critical) {
    for (const pred of t.predecessors) {
      if (!criticalIds.has(pred)) continue;
      lines.push(`  T${pred} --> T${t.uid}`);
    }
  }
  for (const t of critical) {
    lines.push(`  class T${t.uid} critical`);
  }
  lines.push("  classDef critical fill:#fee2e2,stroke:#dc2626,stroke-width:2px,color:#7f1d1d");
  return lines.join("\n");
}

function renderHtml(data) {
  const safeDataJson = JSON.stringify(data)
    .replace(/</g, "\\u003c")
    .replace(/>/g, "\\u003e")
    .replace(/&/g, "\\u0026")
    .replace(/\u2028/g, "\\u2028")
    .replace(/\u2029/g, "\\u2029");
  return `<!doctype html>
<html lang="en">
<head>
<meta charset="UTF-8"/>
<meta name="viewport" content="width=device-width,initial-scale=1.0"/>
<title>Ametek Schedule Intelligence</title>
<style>
*{box-sizing:border-box;margin:0;padding:0}
:root{
  --bg:#f0f4f8;--card:#fff;--line:#dde3ed;--txt:#0f2137;--muted:#5c6f85;
  --red:#c0392b;--red-bg:#fef2f2;--red-bd:#fca5a5;
  --amber:#b45309;--amber-bg:#fffbeb;--amber-bd:#fcd34d;
  --blue:#1d4ed8;--blue-bg:#eff6ff;--blue-bd:#93c5fd;
  --green:#15803d;--green-bg:#f0fdf4;--green-bd:#86efac;
  --ghost:#d1d5db;--today:#dc2626;
  --header:#0b1220;--shadow:0 4px 16px rgba(15,23,42,.09)
}
body{font:13px/1.5 "Segoe UI",system-ui,sans-serif;background:var(--bg);color:var(--txt);display:flex;flex-direction:column;height:100vh;overflow:hidden}

/* HEADER */
header{background:var(--header);color:#fff;padding:10px 18px;display:flex;justify-content:space-between;align-items:center;gap:12px;flex-shrink:0}
header h1{font-size:16px;font-weight:600;letter-spacing:-.3px}
header .sub{font-size:11px;opacity:.7;margin-top:2px}
#metaLine{font-size:11px;opacity:.6;text-align:right}

/* KPI STRIP */
#kpiStrip{display:flex;flex-direction:column;flex-shrink:0;border-bottom:2px solid var(--line);background:var(--card)}
#kpiStrip .strip-label{font-size:9px;font-weight:700;text-transform:uppercase;letter-spacing:.6px;color:var(--muted);padding:4px 14px 0;border-bottom:1px solid var(--line)}
#kpiStrip .strip-label span{display:inline-block;padding:1px 7px;border-radius:9px;font-size:9px;font-weight:700;margin-left:6px}
.strip-cpm{background:#dbeafe;color:#1e40af}.strip-derived{background:#fef9c3;color:#92400e}
.kpi-row{display:flex;gap:0}
.kpi-cell{flex:1;padding:4px 8px;text-align:center;border-right:1px solid var(--line)}
.kpi-cell:last-child{border-right:none}
.kpi-cell .val{font-size:14px;font-weight:700;line-height:1}
.kpi-cell .lbl{font-size:9px;text-transform:uppercase;letter-spacing:.35px;color:var(--muted);margin-top:1px}
.val-red{color:var(--red)}.val-amber{color:var(--amber)}.val-blue{color:var(--blue)}.val-green{color:var(--green)}.val-dim{color:var(--muted)}.val-indigo{color:#4338ca}

#supplementControls{display:flex;align-items:center;gap:8px;padding:6px 12px;background:#f8fafc;border-bottom:1px solid var(--line);flex-shrink:0}
#supplementControls .muted{font-size:10px;color:var(--muted);font-weight:600;text-transform:uppercase;letter-spacing:.35px}
#supplementControls button{font-size:11px;padding:4px 10px;border:1px solid var(--line);border-radius:999px;background:#fff;color:var(--muted);cursor:pointer}
#supplementControls button.active{background:#1e3a8a;color:#fff;border-color:#1e3a8a}
#supplementControls .layout-btn.active{background:#0f766e;border-color:#0f766e}

/* DECISION STRIP */
#decisionStrip{display:grid;grid-template-columns:repeat(4,minmax(180px,1fr));gap:10px;padding:10px 12px;background:var(--bg);border-bottom:1px solid var(--line);flex-shrink:0}
.decision-card{background:var(--card);border:1px solid var(--line);border-radius:10px;padding:10px 12px;box-shadow:var(--shadow);min-height:92px;cursor:pointer;transition:all 150ms ease}
.decision-card.clickable{cursor:pointer}
.decision-card.clickable:hover{border-color:#94a3b8;box-shadow:0 6px 18px rgba(15,23,42,.14);transform:translateY(-2px)}
.decision-card .eyebrow{font-size:10px;font-weight:700;letter-spacing:.45px;text-transform:uppercase;color:var(--muted);margin-bottom:5px}
.decision-card .headline{font-size:13px;font-weight:700;line-height:1.3;color:var(--txt)}
.decision-card .subline{font-size:11px;color:var(--muted);margin-top:4px;line-height:1.4}
.decision-card .pill{display:inline-block;margin-top:6px;padding:2px 8px;border-radius:999px;font-size:10px;font-weight:700}
.pill-red{background:var(--red-bg);color:var(--red)}.pill-amber{background:var(--amber-bg);color:var(--amber)}.pill-blue{background:var(--blue-bg);color:var(--blue)}.pill-green{background:var(--green-bg);color:var(--green)}

/* CPM WALKTHROUGH */
#cpmExplain{display:grid;grid-template-columns:repeat(4,minmax(220px,1fr));gap:10px;padding:10px 12px;background:#eef2ff;border-bottom:1px solid var(--line);flex-shrink:0}
.cpm-card{background:#fff;border:1px solid #c7d2fe;border-radius:10px;padding:10px 12px;box-shadow:var(--shadow);min-height:138px;cursor:pointer;transition:all 150ms ease}
.cpm-card.clickable{cursor:pointer}
.cpm-card.clickable:hover{border-color:#6366f1;box-shadow:0 6px 18px rgba(79,70,229,.15);transform:translateY(-2px)}
.cpm-card .eyebrow{font-size:10px;font-weight:700;letter-spacing:.45px;text-transform:uppercase;color:#3730a3;margin-bottom:5px}
.cpm-card .title{font-size:13px;font-weight:700;line-height:1.3;color:var(--txt)}
.cpm-card .meta{font-size:11px;color:var(--muted);margin-top:4px;line-height:1.4}
.cpm-card .row{display:grid;grid-template-columns:58px 1fr;gap:2px 6px;margin-top:5px;font-size:11px}
.cpm-card .lbl{font-weight:700;color:#4f46e5;text-transform:uppercase;font-size:10px;letter-spacing:.25px}
.cpm-empty{padding:12px;border:1px dashed #a5b4fc;border-radius:8px;background:#fff;color:#4c1d95;font-size:12px}

/* MILESTONE COMMAND CENTER */
#milestoneCommandCenter{padding:10px 12px;background:#eef2ff;border-bottom:1px solid var(--line);flex-shrink:0}
.cmd-wrap{background:#fff;border:1px solid #c7d2fe;border-radius:10px;overflow:hidden;box-shadow:var(--shadow)}
.cmd-head{display:flex;justify-content:space-between;align-items:center;padding:8px 10px;background:#f8faff;border-bottom:1px solid #dbe4ff}
.cmd-title{font-size:12px;font-weight:700;color:#1e3a8a;text-transform:uppercase;letter-spacing:.35px}
.cmd-sub{font-size:11px;color:#5c6f85}
.cmd-table-wrap{max-height:220px;overflow:auto}
.cmd-table{width:100%;border-collapse:collapse;font-size:11px}
.cmd-table th,.cmd-table td{padding:7px 8px;border-bottom:1px solid #eef2ff;text-align:left;vertical-align:middle}
.cmd-table th{position:sticky;top:0;background:#f8faff;color:#334155;font-size:10px;text-transform:uppercase;letter-spacing:.3px;z-index:2}
.cmd-table tr:hover{background:#f8fafc}
.cmd-badge{display:inline-block;padding:2px 8px;border-radius:999px;font-weight:700;font-size:10px}
.cmd-badge.critical{background:#fef2f2;color:#b91c1c}
.cmd-badge.offtrack{background:#fff7ed;color:#c2410c}
.cmd-badge.slipping{background:#fffbeb;color:#a16207}
.cmd-badge.ontrack{background:#f0fdf4;color:#166534}
.cmd-impact{font-weight:700;color:#7f1d1d}

/* TOP OVERVIEW + MODES */
#topOverview{padding:10px 12px;background:#f8fafc;border-bottom:1px solid var(--line);flex-shrink:0}
.overview-card{background:#fff;border:1px solid #dbe4ff;border-radius:10px;box-shadow:var(--shadow);padding:8px 10px}
.overview-head{display:flex;justify-content:space-between;align-items:center;gap:10px;margin-bottom:6px}
.overview-title{font-size:12px;font-weight:700;color:#1e3a8a;text-transform:uppercase;letter-spacing:.35px}
.overview-subtabs{display:flex;gap:6px}
.overview-btn{font-size:11px;font-weight:700;padding:4px 10px;border:1px solid #cbd5e1;border-radius:999px;background:#fff;color:#334155;cursor:pointer}
.overview-btn.active{background:#1e3a8a;color:#fff;border-color:#1e3a8a}
#vegaMilestoneOverview,#visMilestoneOverview{min-height:220px}
#visMilestoneOverview{display:none;border:1px solid #e2e8f0;border-radius:8px;overflow:hidden}
#visMilestoneOverview .vis-item{border-width:1px;border-radius:10px;font-size:11px}
#visMilestoneOverview .vis-item.bar-critical{background:#fee2e2;border-color:#b91c1c;color:#7f1d1d}
#visMilestoneOverview .vis-item.bar-immediate{background:#fef2f2;border-color:#c0392b;color:#7f1d1d}
#visMilestoneOverview .vis-item.bar-soon{background:#fffbeb;border-color:#d97706;color:#92400e}
#visMilestoneOverview .vis-item.bar-watch{background:#eff6ff;border-color:#2563eb;color:#1e40af}
#visMilestoneOverview .vis-group{font-size:11px;font-weight:700;color:#334155}
#visMilestoneOverview .vis-time-axis .vis-text{font-size:10px;color:#475569}
#modeTabs{display:flex;gap:8px;padding:8px 12px;background:#f8fafc;border-bottom:1px solid var(--line);flex-shrink:0}
.mode-btn{font-size:12px;font-weight:700;padding:6px 12px;border:1px solid #cbd5e1;border-radius:6px;background:#fff;color:#334155;cursor:pointer;transition:all 120ms ease}
.mode-btn.active{background:#0f2137;color:#fff;border-color:#0f2137}

body.minimal-ui #supplementControls,
body.minimal-ui #decisionStrip,
body.minimal-ui #cpmExplain,
body.minimal-ui #criticalChainStrip,
body.minimal-ui #wsStrip,
body.minimal-ui #triage,
body.minimal-ui #cycleChips,
body.minimal-ui #trackChips,
body.minimal-ui #filterTier,
body.minimal-ui #filterWs,
body.minimal-ui #filterLv,
body.minimal-ui #focusMilestone,
body.minimal-ui #btnFocusDrivers,
body.minimal-ui #btnContext,
body.minimal-ui #btnCriticalNetwork,
body.minimal-ui .ctrl-section,
body.minimal-ui .ctrl-sep{display:none !important}
body.minimal-ui #ganttControls{display:flex;align-items:center;gap:8px;padding:8px 12px}
body.minimal-ui #ganttInfo{margin-left:auto}
body.minimal-ui #layout{min-width:1080px}

body.simple-mode #decisionStrip,
body.simple-mode #cpmExplain,
body.simple-mode #criticalChainStrip,
body.simple-mode #wsStrip{display:none !important}
body.simple-mode #layout{min-width:1080px}
body.simple-mode #ganttControls{padding:8px 10px;gap:8px}
body.simple-mode #ganttPanel{min-width:980px}
body.simple-mode #triage{display:none !important}

/* CRITICAL CHAIN & TRIAGE */
#criticalChainStrip{display:grid;grid-template-columns:1fr 1fr;gap:10px;padding:12px;background:#fee2e2;border-bottom:1px solid var(--line);flex-shrink:0}
.chain-panel,.triage-panel{background:#fff;border:1px solid #fca5a5;border-radius:10px;padding:12px;box-shadow:var(--shadow)}
.chain-panel .title,.triage-panel .title{font-size:12px;font-weight:700;text-transform:uppercase;letter-spacing:.4px;color:#7f1d1d;margin-bottom:8px;display:flex;align-items:center;gap:6px}
.chain-sequence{display:flex;align-items:center;gap:6px;flex-wrap:wrap;margin-bottom:8px;font-size:11px}
.chain-task{display:inline-block;padding:4px 8px;border-radius:6px;background:#fff5f5;border:1px solid #fecaca;color:#7f1d1d;font-weight:600;max-width:140px;overflow:hidden;text-overflow:ellipsis;white-space:nowrap}
.chain-task.at-risk{background:#fef2f2;border-color:#dc2626;color:#b91c1c;font-weight:700}
.chain-sep{display:inline-block;color:#94a3b8;margin:0 2px}
.chain-stats{font-size:10px;color:#5c6f85;line-height:1.4;padding:6px 0;border-top:1px solid #fecaca;margin-top:6px}
.chain-stat-row{display:flex;justify-content:space-between;margin-bottom:2px}
.chain-stat-val{font-weight:600;color:#7f1d1d}
.triage-list{display:flex;flex-direction:column;gap:6px;font-size:11px;max-height:140px;overflow-y:auto}

/* TASK HIERARCHY TREE */
#taskHierarchy{font-size:11px;color:#334155}
#taskHierarchy .tree-item{margin-left:0;padding:4px 6px;cursor:pointer;border-radius:4px;user-select:none}
#taskHierarchy .tree-item:hover{background:#f1f5f9}
#taskHierarchy .tree-item.selected{background:#0f2137;color:#fff}
#taskHierarchy .tree-item.lv2{font-weight:700;margin-top:6px;color:#1e3a8a}
#taskHierarchy .tree-item.lv3{margin-left:16px;color:#475569;font-size:10px}
#taskHierarchy .tree-toggle{display:inline-block;width:12px;text-align:center;cursor:pointer;margin-right:4px}
#taskHierarchy .tree-toggle.collapsed::before{content:'▶'}
#taskHierarchy .tree-toggle.expanded::before{content:'▼'}
#taskHierarchy .tree-children{display:block;margin-left:8px}
#taskHierarchy .tree-children.hidden{display:none}
.triage-item{padding:6px 8px;border-radius:6px;border-left:3px solid var(--line);background:#f9fafb;display:grid;grid-template-columns:24px 1fr 40px;gap:6px;align-items:center}
.triage-item.tier-1{border-left-color:#dc2626;background:#fef2f2}
.triage-item.tier-2{border-left-color:#f97316;background:#fff7ed}
.triage-item.tier-1 .tier-icon{color:#dc2626;font-weight:700}
.triage-item.tier-2 .tier-icon{color:#f97316;font-weight:700}
.triage-name{font-weight:600;color:#0f2137;overflow:hidden;text-overflow:ellipsis;white-space:nowrap;max-width:180px}
.triage-slip{text-align:right;color:#7f1d1d;font-weight:600;font-size:10px}

/* MAIN LAYOUT */
#layout{display:flex;flex-direction:column;flex:1;overflow:hidden;gap:0;min-height:0}
#layoutTop{display:flex;gap:0;flex:1;overflow:hidden;min-height:0}
#taskHierarchy{flex:0 0 280px;border-right:1px solid var(--line);background:#fff;overflow-y:auto;padding:10px}
#ganttPanelTop{flex:1;display:flex;flex-direction:column;overflow:hidden}
#layoutBottom{display:flex;flex:0 0 320px;border-top:1px solid var(--line);gap:0;overflow:hidden}
#networkSection{flex:1;border-right:1px solid var(--line);background:#fff;padding:10px;overflow-y:auto}
#actionItemsSection{flex:1;background:#fff;padding:10px;overflow-y:auto}
#layout.triage-collapsed #triage{flex:0 0 0;min-width:0;max-width:0;border-right:none;overflow:hidden}
#layout.triage-collapsed #ganttPanel{min-width:1180px}

/* LEFT PANEL — Triage */
#triage{flex:0 0 30%;min-width:320px;max-width:460px;display:flex;flex-direction:column;border-right:2px solid var(--line);background:var(--card)}
#triageTabs{display:flex;flex-shrink:0;border-bottom:2px solid var(--line)}
.tab-btn{flex:1;padding:8px 6px;font-size:12px;font-weight:600;text-transform:uppercase;letter-spacing:.4px;cursor:pointer;border:none;background:transparent;color:var(--muted);border-bottom:3px solid transparent;transition:all .15s}
.tab-btn.active-red{color:var(--red);border-bottom-color:var(--red);background:var(--red-bg)}
.tab-btn.active-amber{color:var(--amber);border-bottom-color:var(--amber);background:var(--amber-bg)}
.tab-btn.active-blue{color:var(--blue);border-bottom-color:var(--blue);background:var(--blue-bg)}
.tab-btn .badge{display:inline-block;margin-left:4px;padding:1px 6px;border-radius:9px;font-size:10px;color:#fff}
.badge-red{background:var(--red)}.badge-amber{background:var(--amber)}.badge-blue{background:var(--blue)}
#triageCards{flex:1;overflow-y:auto;padding:10px}

/* TASK CARD */
.task-card{border-radius:8px;border:1px solid var(--line);padding:10px 12px;margin-bottom:8px;cursor:pointer;transition:box-shadow .12s}
.task-card:hover{box-shadow:0 2px 8px rgba(0,0,0,.12)}
.task-card.tier-IMMEDIATE{border-left:4px solid var(--red);background:var(--red-bg)}
.task-card.tier-SOON{border-left:4px solid var(--amber);background:var(--amber-bg)}
.task-card.tier-WATCH{border-left:4px solid var(--blue);background:var(--blue-bg)}
.card-header{display:flex;justify-content:space-between;align-items:flex-start;gap:6px;margin-bottom:6px}
.card-name{font-size:12px;font-weight:600;color:var(--txt);line-height:1.3}
.card-meta{font-size:10px;color:var(--muted);white-space:nowrap}
.card-wbs{font-size:10px;color:var(--muted);margin-bottom:4px}
.tier-badge{font-size:9px;font-weight:700;padding:2px 7px;border-radius:9px;color:#fff;white-space:nowrap}
.tier-IMMEDIATE .tier-badge{background:var(--red)}.tier-SOON .tier-badge{background:var(--amber)}.tier-WATCH .tier-badge{background:var(--blue)}
.narrative{margin-top:6px;display:none}
.task-card.expanded .narrative{display:block}
.nar-row{display:grid;grid-template-columns:80px 1fr;gap:2px 6px;margin-bottom:3px;font-size:11px}
.nar-label{font-weight:600;color:var(--muted);text-transform:uppercase;font-size:10px;letter-spacing:.3px}
.nar-val{color:var(--txt)}

/* RIGHT PANEL — Gantt */
#ganttPanel{flex:1 1 auto;min-width:840px;display:flex;flex-direction:column;overflow:hidden;background:var(--bg)}
#ganttControls{display:grid;gap:12px;padding:12px;background:var(--bg-alt);font-size:12px;grid-template-columns:repeat(auto-fit, minmax(180px, 1fr))}
#ganttControls label{font-size:11px;color:#666;font-weight:500;margin:0;display:block;margin-bottom:4px}
#ganttControls select{font-size:12px;padding:4px 8px;border:1px solid var(--line);border-radius:6px;background:#fff;color:var(--txt);width:100%;box-sizing:border-box}
#ganttControls button{font-size:11px;padding:4px 10px;border:1px solid var(--line);border-radius:6px;background:#fff;cursor:pointer;color:var(--txt);transition:all 150ms ease}
#ganttControls button:hover{background:var(--bg);border-color:var(--primary)}
.ctrl-section{display:flex;flex-direction:column;gap:6px;padding:8px;background:#fff;border-radius:4px;border:1px solid var(--line)}
.ctrl-section-title{font-size:11px;font-weight:600;color:#333;margin-bottom:2px;text-transform:uppercase;letter-spacing:0.5px}
.ctrl-sep{grid-column:1/-1;height:1px;background:var(--line);margin:4px 0}
#ganttWrap{flex:1;overflow:auto;background:#fff;min-height:420px}
.chip-group{display:flex;gap:6px;flex-wrap:wrap;align-items:center}
.filter-chip{font-size:11px;padding:4px 10px;border:1px solid var(--line);border-radius:999px;background:#fff;color:var(--muted);cursor:pointer;transition:all .12s}
.filter-chip:hover{border-color:#94a3b8;color:var(--txt)}
.filter-chip.active{background:#0f2137;color:#fff;border-color:#0f2137}
.filter-chip.track.active{background:#1e3a8a;border-color:#1e3a8a}
.filter-chip.toggle{width:100%;text-align:center;padding:4px 8px;background:#f9f9f9;border-color:#ddd}
.filter-chip.toggle.active{background:#ff6b6b;color:#fff;border-color:#ff6b6b}
.filter-chip.toggle:hover{background:#f0f0f0}
.filter-chip.toggle.active:hover{background:#ff5252}

/* Frappe Gantt overrides — MS Project visual style */
.gantt-container{background:#fff}
.gantt .grid-background{fill:#fff}
.gantt .grid-header{fill:#f1f5f9;stroke:#dde3ed;stroke-width:1}
.gantt .grid-row{fill:transparent}
.gantt .grid-row:nth-child(even){fill:#f8fafc}
.gantt .row-line{stroke:#edf2f7;stroke-width:1}
.gantt .tick{stroke:#dde3ed;stroke-width:1}
.gantt .tick.thick{stroke:#c7d2e0;stroke-width:2}
.gantt .today-highlight{fill:#dc262615;stroke:none}
.gantt .today-highlight + line,.gantt line.today{stroke:#dc2626;stroke-width:2}
.gantt .upper-text{fill:#374151;font-weight:700;font-size:12px}
.gantt .lower-text{fill:#6b7280;font-size:11px}
.gantt .bar-label{fill:#fff;font-size:11px;font-weight:500}
.gantt .bar-label.big{fill:#374151}
/* Tier-colored bars */
.gantt .bar-wrapper.bar-immediate .bar{fill:#c0392b;stroke:#9b1e17}
.gantt .bar-wrapper.bar-immediate .bar-progress{fill:#922b21}
.gantt .bar-wrapper.bar-immediate .bar-label{fill:#fff}
.gantt .bar-wrapper.bar-soon .bar{fill:#d97706;stroke:#b45309}
.gantt .bar-wrapper.bar-soon .bar-progress{fill:#92400e}
.gantt .bar-wrapper.bar-soon .bar-label{fill:#fff}
.gantt .bar-wrapper.bar-watch .bar{fill:#2563eb;stroke:#1d4ed8}
.gantt .bar-wrapper.bar-watch .bar-progress{fill:#1e3a8a}
.gantt .bar-wrapper.bar-watch .bar-label{fill:#fff}
.gantt .bar-wrapper.bar-summary .bar{fill:#64748b;stroke:#475569}
.gantt .bar-wrapper.bar-summary .bar-progress{fill:#475569}
.gantt .bar-wrapper.bar-summary .bar-label{fill:#fff}
.gantt .bar-wrapper.bar-track .bar{fill:#475569;stroke:#334155}
.gantt .bar-wrapper.bar-track .bar-progress{fill:#334155}
.gantt .bar-wrapper.bar-track .bar-label{fill:#f8fafc;font-weight:700}
.gantt .bar-wrapper.bar-lane .bar{fill:#0f2137;stroke:#0b1220}
.gantt .bar-wrapper.bar-lane .bar-progress{fill:#0f2137}
.gantt .bar-wrapper.bar-lane .bar-label{fill:#fff;font-weight:700;letter-spacing:.3px}
.gantt .bar-wrapper.bar-focus .bar{fill:#6d28d9;stroke:#4c1d95}
.gantt .bar-wrapper.bar-focus .bar-progress{fill:#4c1d95}
.gantt .bar-wrapper.bar-focus .bar-label{fill:#f5f3ff;font-weight:700}
.gantt .bar-wrapper.context-driver .bar{fill:#dcfce7;stroke:#16a34a;stroke-dasharray:0}
.gantt .bar-wrapper.context-driver .bar-progress{fill:#86efac}
.gantt .bar-wrapper.context-driver .bar-label{fill:#166534;font-weight:700}
.gantt .bar-wrapper.context-upstream .bar,.gantt .bar-wrapper.context-downstream .bar{fill:#e2e8f0;stroke:#94a3b8;stroke-dasharray:4 2}
.gantt .bar-wrapper.context-upstream .bar-progress,.gantt .bar-wrapper.context-downstream .bar-progress{fill:#cbd5e1}
.gantt .bar-wrapper.context-upstream .bar-label,.gantt .bar-wrapper.context-downstream .bar-label{fill:#334155}
.gantt .bar-wrapper.bar-critical .bar{fill:#7f1d1d;stroke:#991b1b}
.gantt .bar-wrapper.bar-critical .bar-progress{fill:#450a0a}
.gantt .bar-wrapper.bar-critical .bar-label{fill:#fecaca}
/* MSP-style status bar colors (% complete + on-time vs late) */
.gantt .bar-wrapper.bar-msp-complete .bar{fill:#94a3b8;stroke:#64748b}
.gantt .bar-wrapper.bar-msp-complete .bar-progress{fill:#64748b}
.gantt .bar-wrapper.bar-msp-complete .bar-label{fill:#fff}
.gantt .bar-wrapper.bar-msp-late .bar{fill:#dc2626;stroke:#b91c1c}
.gantt .bar-wrapper.bar-msp-late .bar-progress{fill:#991b1b}
.gantt .bar-wrapper.bar-msp-late .bar-label{fill:#fff}
.gantt .bar-wrapper.bar-msp-inprogress .bar{fill:#1d4ed8;stroke:#1e40af}
.gantt .bar-wrapper.bar-msp-inprogress .bar-progress{fill:#1e3a8a}
.gantt .bar-wrapper.bar-msp-inprogress .bar-label{fill:#fff}
.gantt .bar-wrapper.bar-msp-notstarted .bar{fill:#3b82f6;stroke:#2563eb}
.gantt .bar-wrapper.bar-msp-notstarted .bar-progress{fill:#1d4ed8}
.gantt .bar-wrapper.bar-msp-notstarted .bar-label{fill:#fff}
/* Swimlane band rects injected into SVG after render */
.sl-band{pointer-events:none}
/* Arrow connectors */
.gantt .arrow{stroke:#94a3b8;stroke-width:1.5;fill:none}
.gantt .arrow path{stroke:#94a3b8}
/* hover - interactive affordance */
.gantt .bar-wrapper .bar{cursor:pointer;transition:opacity 100ms ease}
.gantt .bar-wrapper:hover .bar{opacity:0.9}
.gantt .bar-wrapper.active .bar{stroke-width:2;stroke:#0f2137}

/* Popup */
.gantt-popup{font:12px/1.5 "Segoe UI",system-ui,sans-serif;min-width:280px;max-width:360px}
.gantt-popup .title{font-weight:700;font-size:13px;color:#0f2137;margin-bottom:6px;padding-bottom:6px;border-bottom:1px solid #dde3ed}
.gantt-popup .popup-row{display:grid;grid-template-columns:72px 1fr;gap:2px 6px;margin-bottom:2px;font-size:11px}
.gantt-popup .popup-lbl{font-weight:600;color:#5c6f85;text-transform:uppercase;font-size:10px}
.pop-imm{color:#c0392b;font-weight:700}.pop-soon{color:#b45309;font-weight:700}.pop-watch{color:#1d4ed8;font-weight:700}

/* BOTTOM WORKSTREAM STRIP */
#wsStrip{flex-shrink:0;background:var(--card);border-top:2px solid var(--line);padding:8px 12px;display:flex;gap:6px;overflow-x:auto;max-height:96px}
.ws-chip{flex-shrink:0;border:1px solid var(--line);border-radius:8px;padding:6px 10px;cursor:pointer;min-width:130px;transition:box-shadow .12s}
.ws-chip:hover{box-shadow:0 2px 6px rgba(0,0,0,.10)}
.ws-chip.active{border-color:var(--blue);box-shadow:0 0 0 2px var(--blue-bd)}
.ws-name{font-size:11px;font-weight:600;color:var(--txt);white-space:nowrap;overflow:hidden;text-overflow:ellipsis}
.ws-bars{display:flex;height:6px;border-radius:3px;overflow:hidden;margin-top:4px}
.ws-bars span{display:block;height:100%}
.ws-counts{font-size:10px;color:var(--muted);margin-top:3px}

/* tooltip */
#tooltip{position:fixed;background:#1e293b;color:#e2e8f0;font-size:11px;padding:8px 12px;border-radius:8px;pointer-events:none;z-index:999;max-width:320px;display:none;line-height:1.55}

/* ACTION PANEL */
#apToggleBtn{position:fixed;right:0;top:50%;transform:translateY(-50%);z-index:850;writing-mode:vertical-rl;text-orientation:mixed;padding:12px 6px;background:#1e3a8a;color:#fff;border:none;border-radius:8px 0 0 8px;font-size:11px;font-weight:700;cursor:pointer;letter-spacing:.5px;box-shadow:-2px 0 8px rgba(0,0,0,.18);transition:background .15s}
#apToggleBtn:hover{background:#1e40af}
#apToggleBtn .ap-count-dot{display:inline-block;background:#dc2626;color:#fff;border-radius:999px;font-size:10px;font-weight:700;padding:1px 5px;margin-top:4px;writing-mode:horizontal-tb}
#actionPanel{position:fixed;top:0;right:0;bottom:0;width:25vw;min-width:280px;max-width:420px;background:#fff;border-left:2px solid #e2e8f0;box-shadow:-4px 0 24px rgba(15,23,42,.12);z-index:840;display:flex;flex-direction:column;transform:translateX(100%);transition:transform 0.3s ease}
#actionPanel.open{transform:translateX(0)}
.ap-header{display:flex;align-items:center;gap:8px;padding:10px 12px;background:#1e3a8a;color:#fff;flex-shrink:0}
.ap-title{font-size:13px;font-weight:700;flex:1}
.ap-count{font-size:11px;background:rgba(255,255,255,.25);padding:2px 8px;border-radius:999px}
.ap-header button{background:transparent;border:1px solid rgba(255,255,255,.4);color:#fff;border-radius:6px;padding:3px 8px;font-size:11px;cursor:pointer}
.ap-sub{font-size:10px;color:#64748b;padding:6px 12px;background:#f8fafc;border-bottom:1px solid #e2e8f0;flex-shrink:0}
.ap-lanes{display:flex;gap:4px;padding:6px 10px;background:#f8fafc;border-bottom:1px solid #e2e8f0;flex-shrink:0;flex-wrap:wrap}
.ap-lane-chip{font-size:10px;font-weight:700;padding:3px 9px;border-radius:999px;border:1px solid #e2e8f0;background:#fff;color:#64748b;cursor:pointer;transition:all .12s}
.ap-lane-chip.active{background:#1e3a8a;color:#fff;border-color:#1e3a8a}
.ap-lane-chip.lane-now.active{background:#dc2626;border-color:#dc2626}
.ap-lane-chip.lane-risk.active{background:#d97706;border-color:#d97706}
.ap-lane-chip.lane-watch.active{background:#2563eb;border-color:#2563eb}
#apList{flex:1;overflow-y:auto;padding:8px}
.ap-item{border:1px solid #e2e8f0;border-radius:8px;padding:9px 10px;margin-bottom:6px;background:#fff;cursor:default;transition:box-shadow .1s}
.ap-item:hover{box-shadow:0 2px 8px rgba(0,0,0,.08)}
.ap-item.lane-now{border-left:3px solid #dc2626}
.ap-item.lane-risk{border-left:3px solid #d97706}
.ap-item.lane-watch{border-left:3px solid #2563eb}
.ap-item-top{display:flex;align-items:center;gap:6px;margin-bottom:4px}
.ap-lane-badge{font-size:9px;font-weight:700;padding:1px 7px;border-radius:999px}
.ap-lane-badge.now{background:#fef2f2;color:#b91c1c}
.ap-lane-badge.risk{background:#fff7ed;color:#c2410c}
.ap-lane-badge.watch{background:#eff6ff;color:#1d4ed8}
.ap-sev{font-size:9px;padding:1px 7px;border-radius:999px;font-weight:700;margin-left:auto}
.ap-sev.critical{background:#fef2f2;color:#b91c1c}
.ap-sev.high{background:#fff7ed;color:#c2410c}
.ap-sev.medium{background:#fffbeb;color:#a16207}
.ap-sev.low{background:#f0fdf4;color:#166534}
.ap-item-name{font-size:11px;font-weight:600;color:#0f172a;margin-bottom:3px;line-height:1.35;overflow:hidden;text-overflow:ellipsis;display:-webkit-box;-webkit-line-clamp:2;-webkit-box-orient:vertical}
.ap-item-task{font-size:10px;color:#475569;white-space:nowrap;overflow:hidden;text-overflow:ellipsis;margin-bottom:2px}
.ap-item-meta{font-size:10px;color:#94a3b8}
.ap-empty{padding:24px 16px;text-align:center;color:#94a3b8;font-size:12px}
/* NETWORK CANVAS DRAWER */
#networkCanvas{position:fixed;bottom:0;left:0;right:0;height:0;transition:height 0.35s ease;background:#fff;border-top:2px solid #cbd5e1;box-shadow:0 -4px 24px rgba(0,0,0,.15);z-index:900;display:flex;flex-direction:column;overflow:hidden}
#networkCanvas.open{height:44vh}
.nc-handle{display:flex;align-items:center;gap:10px;padding:6px 12px;background:#f8fafc;border-bottom:1px solid #e2e8f0;flex-shrink:0}
.nc-drag-bar{width:40px;height:4px;background:#cbd5e1;border-radius:2px;margin-right:4px}
.nc-mode-chip{font-size:11px;font-weight:700;color:#1e40af;background:#dbeafe;padding:2px 10px;border-radius:999px}
.nc-handle button{margin-left:auto;font-size:11px;padding:3px 10px;border:1px solid #cbd5e1;border-radius:8px;background:#fff;cursor:pointer}
.nc-milestone-strip{display:flex;gap:8px;padding:8px 12px;overflow-x:auto;flex-shrink:0;border-bottom:1px solid #e2e8f0;min-height:0}
.nc-mcard{background:#f8fafc;border:2px solid #e2e8f0;border-radius:8px;padding:8px 10px;min-width:180px;max-width:220px;flex-shrink:0;cursor:pointer}
.nc-mcard.cp-card{border-color:#dc2626}.nc-mcard.driver-card{border-color:#1d4ed8}
.nc-mcard-top{display:flex;justify-content:space-between;align-items:center;margin-bottom:3px}
.nc-mcard-ws{font-size:10px;color:#64748b;font-weight:600;max-width:100px;overflow:hidden;text-overflow:ellipsis;white-space:nowrap}
.nc-mcard-sf{font-size:10px;color:#64748b}
.nc-mcard-badges{display:flex;gap:4px;flex-wrap:wrap;margin-bottom:3px}
.nc-mcard-name{font-size:12px;font-weight:600;white-space:nowrap;overflow:hidden;text-overflow:ellipsis;color:#0f172a}
.nc-graph-host{flex:1;min-height:0}
.nc-callouts{display:flex;gap:10px;padding:8px 12px;flex-wrap:wrap;border-top:1px solid #e2e8f0;flex-shrink:0;overflow-y:auto;max-height:100px}
.nc-callout{background:#f1f5f9;border-radius:6px;padding:6px 10px;font-size:11px;max-width:300px;line-height:1.45}
.nc-callout strong{display:block;font-size:10px;text-transform:uppercase;letter-spacing:.3px;color:#475569;margin-bottom:3px}
/* NETWORK MODAL */
#networkModal{position:fixed;inset:0;background:rgba(15,23,42,.42);display:none;align-items:center;justify-content:center;z-index:1100}
#networkModal.open{display:flex}
.network-panel{width:min(980px,94vw);max-height:82vh;overflow:auto;background:#fff;border:1px solid #cbd5e1;border-radius:12px;box-shadow:0 20px 60px rgba(15,23,42,.35)}
.network-hd{display:flex;justify-content:space-between;align-items:center;padding:10px 12px;border-bottom:1px solid #e2e8f0;background:#f8fafc}
.network-title{font-size:14px;font-weight:700;color:#0f172a}
.network-sub{font-size:11px;color:#64748b}
.network-hd button{font-size:11px;padding:4px 10px;border:1px solid #cbd5e1;border-radius:8px;background:#fff;cursor:pointer}
.network-body{padding:12px;display:grid;grid-template-columns:1fr 1fr;gap:12px}
.network-box{border:1px solid #e2e8f0;border-radius:10px;padding:10px;background:#fff}
.network-box h4{font-size:11px;text-transform:uppercase;letter-spacing:.35px;color:#475569;margin-bottom:6px}
.network-list{font-size:12px;line-height:1.45;color:#0f172a}
.network-actions{display:flex;gap:8px;padding:10px 12px;border-top:1px solid #e2e8f0;background:#f8fafc}
.network-actions button{font-size:11px;padding:5px 10px;border:1px solid #cbd5e1;border-radius:8px;background:#fff;cursor:pointer}
</style>
<link rel="stylesheet" href="https://cdn.jsdelivr.net/npm/frappe-gantt/dist/frappe-gantt.css"/>
<link rel="stylesheet" href="https://unpkg.com/vis-timeline/styles/vis-timeline-graph2d.min.css"/>
</head>
<body>
<header>
  <div>
    <h1>Ametek SAP S/4 — Schedule Intelligence</h1>
    <div class="sub">Priority-tiered risk dashboard · Working tasks only · Real-time CPM analysis</div>
  </div>
  <div id="metaLine"></div>
</header>

<div id="kpiStrip">
  <div class="strip-label">
    <span class="strip-cpm">CPM Engine</span> — direct output from critical path analysis
    <span style="margin-left:24px;opacity:.5">|</span>
    <span class="strip-derived" style="margin-left:12px">Risk Tiers</span> — derived from CPM slack + scheduled start date (tasks starting &gt;4 weeks out excluded unless critical)
    <span style="float:right;padding-right:4px" id="metaLine"></span>
  </div>
  <div class="kpi-row">
    <div class="kpi-cell" title="CPM: tasks with zero total slack — any delay extends project end date">
      <div class="val val-red" id="kCritical">—</div><div class="lbl">Critical Path Tasks</div>
    </div>
    <div class="kpi-cell" title="CPM: tasks with ≤10 days total slack — vulnerable to becoming critical">
      <div class="val val-amber" id="kNearCritical">—</div><div class="lbl">Near-Critical (≤10d slack)</div>
    </div>
    <div class="kpi-cell" title="CPM: total project duration in working days from forward pass">
      <div class="val val-indigo" id="kDuration">—</div><div class="lbl">CPM Duration (days)</div>
    </div>
    <div class="kpi-cell" title="CPM: circular dependencies found — would invalidate CPM results">
      <div class="val val-dim" id="kCyclic">—</div><div class="lbl">Logic Cycles</div>
    </div>
    <div class="kpi-cell" style="border-left:2px solid var(--line)" title="Derived: active/near-start tasks that are past-due, critical, or slipping ≥30 days">
      <div class="val val-red" id="kImmediate">—</div><div class="lbl">Immediate</div>
    </div>
    <div class="kpi-cell" title="Derived: tasks starting within 4 weeks with ≥7-day slip or finishing within 14 days">
      <div class="val val-amber" id="kSoon">—</div><div class="lbl">Soon</div>
    </div>
    <div class="kpi-cell" title="Derived: active tasks needing monitoring — low slack or minor slip">
      <div class="val val-blue" id="kWatch">—</div><div class="lbl">Watch</div>
    </div>
    <div class="kpi-cell" title="Working tasks (not summary, not milestone, not complete, active)">
      <div class="val val-green" id="kWorking">—</div><div class="lbl">Working Tasks</div>
    </div>
  </div>
</div>

<div id="topOverview">
  <div class="overview-card">
    <div class="overview-head">
      <div class="overview-title">Milestones Overview</div>
      <div class="overview-subtabs">
        <button id="overviewVega" class="overview-btn active" onclick="setOverviewRenderer('vega')">Vega-Lite</button>
        <button id="overviewVis" class="overview-btn" onclick="setOverviewRenderer('vis')">vis-timeline</button>
      </div>
    </div>
    <div id="vegaMilestoneOverview"></div>
    <div id="visMilestoneOverview"></div>
  </div>
</div>

<div id="supplementControls">
  <span class="muted">Supplementary insights</span>
  <button id="btnSimpleMode" class="active" onclick="toggleSimpleMode()">Simple mode</button>
  <button id="btnToggleTriage" class="layout-btn" onclick="toggleTriagePanel()">Show triage</button>
  <button id="btnToggleDecision" onclick="toggleSupplement('decision')">Decision cards</button>
  <button id="btnToggleCpm" onclick="toggleSupplement('cpm')">CPM walkthrough</button>
</div>

<div id="milestoneCommandCenter"></div>
<div id="decisionStrip"></div>
<div id="cpmExplain"></div>
<div id="criticalChainStrip"></div>

<div id="layout">
  <!-- TOP: Layout with left sidebar + gantt -->
  <div id="layoutTop">
    <!-- LEFT: Task Hierarchy (lv2/lv3) -->
    <div id="taskHierarchy">
      <div style="font-weight:700;font-size:11px;color:#1e3a8a;margin-bottom:10px;text-transform:uppercase;letter-spacing:.3px">Working Tasks (lv2/lv3)</div>
      <div id="taskTree"></div>
    </div>

    <!-- RIGHT: Gantt Panel -->
    <div id="ganttPanelTop">
    <div id="ganttControls">
      <!-- SECTION: SCOPE (Filter what appears) -->
      <div class="ctrl-section">
        <div class="ctrl-section-title">🔍 Scope</div>
        <label title="Pick which test cycles appear on the timeline">Cycles</label>
        <div id="cycleChips" class="chip-group"></div>
        <label title="Pick which work tracks appear (for example Refresh, Data Loads, Testing)">Tracks</label>
        <div id="trackChips" class="chip-group"></div>
      </div>

      <!-- SECTION: FOCUS (Narrow to specific milestone or chain) -->
      <div class="ctrl-section">
        <div class="ctrl-section-title">🎯 Focus</div>
        <label title="Zoom in on one milestone and related tasks">Focus milestone</label>
        <select id="focusMilestone" onchange="selectFocusMilestone(this.value)" title="Pick one milestone to zoom in on"><option value="">All visible lanes</option></select>
        <button id="btnFocusDrivers" class="filter-chip toggle" onclick="toggleFocusDriversOnly()" title="Show only the tasks directly driving this milestone">Driver only</button>
      </div>

      <!-- SECTION: DISPLAY (Refine visibility) -->
      <div class="ctrl-section">
        <div class="ctrl-section-title">👁️ Display</div>
        <label title="One central filter for what this timeline shows">View filter</label>
        <select id="viewFilter" onchange="setTimelineMode(this.value)" title="Choose what this one central Gantt should show">
          <option value="ALL" selected>All key milestones</option>
          <option value="CRITICAL">Critical path</option>
          <option value="DRIVER">Driver flow</option>
          <option value="SLIPPING">Running late (slipping)</option>
          <option value="PAST_DUE">Past due</option>
          <option value="CRITICAL_TODAY">Needs action now</option>
          <option value="NEAR_CRITICAL">Needs attention soon</option>
          <option value="SOON">Coming up soon</option>
          <option value="WATCH">Watch list</option>
        </select>
        <label title="Filter by urgency (Urgent=red, Soon=amber, Watch list=blue)">Show urgency</label>
        <select id="filterTier" onchange="rebuildGantt()" title="Filter tasks by urgency">
          <option value="">All urgency levels</option>
          <option value="IMMEDIATE">Urgent only</option>
          <option value="SOON">Coming up soon</option>
          <option value="WATCH">Watch list only</option>
        </select>
        <label title="Show only one team or work area">Workstream</label>
        <select id="filterWs" onchange="rebuildGantt()" title="Filter to one team/work area"><option value="">All workstreams</option></select>
        <label title="How much detail to show under each anchor">Depth</label>
        <select id="filterLv" onchange="rebuildGantt()" title="Choose how much detail appears under each anchor">
          <option value="1">Anchor + 1 level</option>
          <option value="2" selected>Anchor + 2 levels</option>
          <option value="99">All descendants</option>
        </select>
      </div>

      <!-- SECTION: VIEW (Timeline and network options) -->
      <div class="ctrl-section">
        <div class="ctrl-section-title">📊 View</div>
        <label title="Change the timeline zoom level">Timeline scale</label>
        <select id="viewMode" onchange="changeView()" title="Choose timeline zoom: week, month, day, and more">
          <option value="Week" selected>Week</option>
          <option value="Month">Month</option>
          <option value="Quarter Day">Quarter Day</option>
          <option value="Half Day">Half Day</option>
          <option value="Day">Day</option>
          <option value="Year">Year</option>
        </select>
        <button id="btnContext" class="filter-chip toggle active" onclick="toggleContextMode()" title="Show related tasks before and after the selected work">Show drivers</button>
        <button id="btnCriticalNetwork" class="filter-chip toggle" onclick="toggleCriticalNetwork()" title="Show only tasks tied to the critical path">Critical network</button>
      </div>

      <div class="ctrl-sep"></div>
      <button onclick="scrollFocus()" title="Jump to the most important milestone">🎯 Focus</button>
      <button onclick="scrollToday()" title="Jump to today's date on the timeline">📅 Today</button>
      <span id="ganttInfo" style="font-size:11px;color:var(--muted);margin-left:auto" title="Quick summary of what is currently shown"></span>
    </div>
    <div id="ganttWrap"><svg id="gantt"></svg></div>
    </div>
  </div>

  <!-- MIDDLE: Mode Tabs -->
  <div id="modeTabs">
    <button id="modeAll" class="mode-btn active" onclick="setTimelineMode('ALL')">All milestones view</button>
    <button id="modeCritical" class="mode-btn" onclick="setTimelineMode('CRITICAL')">Critical path view</button>
    <button id="modeDriver" class="mode-btn" onclick="setTimelineMode('DRIVER')">Driver flow view</button>
  </div>

  <!-- BOTTOM: Network + Action Items -->
  <div id="layoutBottom">
    <!-- LEFT: Network Diagram Section -->
    <div id="networkSection">
      <div style="font-weight:700;font-size:11px;color:#1e3a8a;margin-bottom:8px;text-transform:uppercase;letter-spacing:.3px">Network View</div>
      <div id="networkCanvas">
        <div class="nc-handle">
          <div class="nc-drag-bar"></div>
          <span id="ncModeChip" class="nc-mode-chip">Canvas: All milestones</span>
        </div>
        <div id="ncMilestoneStrip" class="nc-milestone-strip"></div>
        <div id="ncGraph" class="nc-graph-host"></div>
        <div id="ncCallouts" class="nc-callouts"></div>
      </div>
    </div>

    <!-- RIGHT: Action Items Section -->
    <div id="actionItemsSection">
      <div style="font-weight:700;font-size:11px;color:#1e3a8a;margin-bottom:8px;text-transform:uppercase;letter-spacing:.3px">Issues & Risks (217)</div>
      <div id="actionPanel" style="border:none;box-shadow:none;padding:0;max-height:none">
        <div id="actionPanelContent"></div>
    </div>
  </div>
</div>
</div>
</div>

<div id="tooltip"></div>

<div id="networkModal" onclick="if(event.target.id==='networkModal') closeNetworkModal()">
  <div class="network-panel">
    <div class="network-hd">
      <div>
        <div id="networkTitle" class="network-title">Related task map</div>
        <div id="networkSubtitle" class="network-sub"></div>
      </div>
      <button onclick="closeNetworkModal()">Close</button>
    </div>
    <div class="network-body">
      <div class="network-box"><h4>Task summary</h4><div id="networkSummary" class="network-list"></div></div>
      <div class="network-box"><h4>What drives this task (before)</h4><div id="networkDrivers" class="network-list"></div></div>
      <div class="network-box"><h4>What this task affects (after)</h4><div id="networkImpacts" class="network-list"></div></div>
      <div class="network-box"><h4>Schedule story</h4><div id="networkNarrative" class="network-list"></div></div>
    </div>
    <div class="network-actions">
      <button id="networkBtnDriver" onclick="focusTask(window.__networkUid || '', { driverOnly: true }); closeNetworkModal();">Focus driving tasks only</button>
      <button id="networkBtnContext" onclick="focusTask(window.__networkUid || '', { driverOnly: false }); closeNetworkModal();">Focus full related map</button>
    </div>
  </div>
</div>

<script src="https://cdn.jsdelivr.net/npm/frappe-gantt/dist/frappe-gantt.umd.js"></script>
<script src="https://cdn.jsdelivr.net/npm/vega@5"></script>
<script src="https://cdn.jsdelivr.net/npm/vega-lite@5"></script>
<script src="https://cdn.jsdelivr.net/npm/vega-embed@6"></script>
<script src="https://unpkg.com/vis-network/standalone/umd/vis-network.min.js"></script>
<script src="https://unpkg.com/vis-timeline/standalone/umd/vis-timeline-graph2d.min.js"></script>
<script>
const DATA = ${safeDataJson};
const m = DATA.metrics;
const laneAnchors = DATA.laneAnchors || [];

// Auto-derive cycles and tracks from discovered lane anchors
const derivedCycles = [...new Set((laneAnchors || []).map((a) => a.cycle))].sort();
const derivedTracks = [...new Set((laneAnchors || []).map((a) => a.track))].sort();
const cycleOrder = derivedCycles.length ? derivedCycles : ['ITC1','ITC2','ITC3','UAT'];
const trackOrder = derivedTracks.length ? derivedTracks : ['Refresh','Data Loads','Testing'];
const selectedCycles = new Set(cycleOrder);
const selectedTracks = new Set(trackOrder);
const smartScrollDate = (m.decisionMetrics || {}).smartScrollDate || null;
let includeContext = true;
let criticalNetworkOnly = false;
let focusedMilestoneUid = '';
let focusedTaskUid = '';
let focusDriversOnly = false;
let showDecisionStrip = false;
let showCpmStrip = false;
let triageVisible = false;
let simpleMode = true;
let timelineMode = 'ALL';
let overviewRenderer = 'vega';
let visOverviewTimeline = null;
const taskByUid = new Map(DATA.tasks.map((t) => [t.uid, t]));

document.body.classList.add('minimal-ui');

// Build UID → tier map from prioritized tasks
const tierMap = new Map();
const taskNarrativeMap = new Map();
for (const tier of ['immediate','soon','watch']) {
  for (const t of (m[tier] || [])) {
    tierMap.set(t.uid, t.tier);
    taskNarrativeMap.set(t.uid, t);
  }
}

// ── KPI Strip ──────────────────────────────────────────────────────────────
// CPM Engine direct outputs
document.getElementById('kCritical').textContent = m.criticalTasks;
document.getElementById('kNearCritical').textContent = m.nearCriticalTasks;
document.getElementById('kDuration').textContent = m.cpmSummary.projectDuration;
document.getElementById('kCyclic').textContent = m.cpmSummary.cyclic ? 'YES ⚠' : 'None';
if (m.cpmSummary.cyclic) document.getElementById('kCyclic').className = 'val val-red';
// Derived risk tiers
document.getElementById('kImmediate').textContent = m.immediateCount;
document.getElementById('kSoon').textContent = m.soonCount;
document.getElementById('kWatch').textContent = m.watchCount;
document.getElementById('kWorking').textContent = m.workingTasks;
const badgeImmediate = document.getElementById('badgeImmediate'); if (badgeImmediate) badgeImmediate.textContent = m.immediateCount;
const badgeSoon = document.getElementById('badgeSoon'); if (badgeSoon) badgeSoon.textContent = m.soonCount;
const badgeWatch = document.getElementById('badgeWatch'); if (badgeWatch) badgeWatch.textContent = m.watchCount;
document.getElementById('metaLine').innerHTML = DATA.source.file + ' · ' + new Date(DATA.generatedAt).toLocaleString();

// ── Triage Panel ───────────────────────────────────────────────────────────
let activeTier = 'IMMEDIATE';

function esc(s){ return String(s||'').replace(/&/g,'&amp;').replace(/</g,'&lt;').replace(/>/g,'&gt;'); }
function fmtD(v){ if(!v) return '—'; const d=new Date(v); return Number.isNaN(d.getTime())?v:d.toLocaleDateString('en-US',{month:'short',day:'numeric',year:'numeric'}); }
function shortText(value, max=64){ const s=String(value||''); return s.length>max ? s.slice(0,max-1)+'…' : s; }

function setTimelineMode(mode) { actionPanelTaskUid = null; setTimeout(function(){ if(actionPanelOpen) renderActionPanel(); }, 50);
  timelineMode = mode || 'ALL';
  const filterEl = document.getElementById('viewFilter');
  if (filterEl && filterEl.value !== timelineMode) filterEl.value = timelineMode;
  refreshViewFilterTooltip();
  rebuildGantt();
}

function viewFilterHelp(mode) {
  const help = {
    ALL: 'Show the full milestone timeline.',
    CRITICAL: 'Tasks on the critical path — if these slip, the overall end date slips.',
    DRIVER: 'Tasks that are driving or influencing key milestone dates.',
    SLIPPING: 'Tasks finishing later than their baseline date.',
    PAST_DUE: 'Tasks that should already be finished but are not complete.',
    CRITICAL_TODAY: 'Highest-priority items that need attention now.',
    NEAR_CRITICAL: 'Tasks close to becoming critical soon.',
    SOON: 'Tasks starting or finishing in the near term.',
    WATCH: 'Tasks to monitor closely for possible risk.'
  };
  return help[mode] || 'Choose what this one central Gantt should show.';
}

function refreshViewFilterTooltip() {
  const filterEl = document.getElementById('viewFilter');
  if (!filterEl) return;
  filterEl.title = viewFilterHelp(filterEl.value || timelineMode || 'ALL');
}

function refreshControlTooltips() {
  const btnContext = document.getElementById('btnContext');
  if (btnContext) {
    btnContext.title = includeContext
      ? 'Related tasks before and after are currently shown. Click to hide them.'
      : 'Related tasks before and after are hidden. Click to show them.';
  }

  const btnCriticalNetwork = document.getElementById('btnCriticalNetwork');
  if (btnCriticalNetwork) {
    btnCriticalNetwork.title = criticalNetworkOnly
      ? 'Showing only tasks tied to the critical path. Click to return to normal.'
      : 'Show only tasks tied to the critical path.';
  }

  const btnFocusDrivers = document.getElementById('btnFocusDrivers');
  if (btnFocusDrivers) {
    btnFocusDrivers.title = focusDriversOnly
      ? 'Showing only direct driving tasks. Click to show full related map.'
      : 'Show only the tasks directly driving this milestone.';
  }

  const viewMode = document.getElementById('viewMode');
  if (viewMode) {
    viewMode.title = 'Current zoom: ' + viewMode.value + '. Change to zoom in or out on time.';
  }
}

function getOverviewMilestoneRows() {
  return (m.decisionMetrics?.milestoneCommandTable || []).slice(0, 40).map((r, idx) => ({
    id: 'ms_' + idx + '_' + r.uid,
    uid: r.uid,
    parentMilestone: r.parentMilestone || 'Program',
    milestone: shortText(r.name, 42),
    fullMilestone: r.name,
    finish: r.finish,
    slipDays: r.slipDays || 0,
    status: r.status || 'On-track'
  })).filter((r) => r.finish);
}

function renderOverviewTabs() {
  const vBtn = document.getElementById('overviewVega');
  const tBtn = document.getElementById('overviewVis');
  if (!vBtn || !tBtn) return;
  vBtn.className = 'overview-btn' + (overviewRenderer === 'vega' ? ' active' : '');
  tBtn.className = 'overview-btn' + (overviewRenderer === 'vis' ? ' active' : '');
}

function setOverviewRenderer(kind) {
  overviewRenderer = kind === 'vis' ? 'vis' : 'vega';
  const vegaEl = document.getElementById('vegaMilestoneOverview');
  const visEl = document.getElementById('visMilestoneOverview');
  if (vegaEl) vegaEl.style.display = overviewRenderer === 'vega' ? 'block' : 'none';
  if (visEl) visEl.style.display = overviewRenderer === 'vis' ? 'block' : 'none';
  renderOverviewTabs();
}

function renderVegaMilestoneOverview() {
  const target = document.getElementById('vegaMilestoneOverview');
  const rows = getOverviewMilestoneRows();

  if (!target) return;
  if (!rows.length) {
    target.innerHTML = '<div style="font-size:12px;color:#64748b;padding:10px">No milestone data available.</div>';
    return;
  }
  if (typeof vegaEmbed !== 'function') {
    target.innerHTML = '<div style="font-size:12px;color:#64748b;padding:10px">Vega-Lite failed to load.</div>';
    return;
  }

  const spec = {
    $schema: 'https://vega.github.io/schema/vega-lite/v5.json',
    width: 'container',
    height: 200,
    data: { values: rows },
    mark: { type: 'point', filled: true, size: 130, shape: 'diamond' },
    encoding: {
      x: { field: 'finish', type: 'temporal', title: 'Finish date' },
      y: { field: 'parentMilestone', type: 'nominal', sort: '-x', title: 'Parent milestone' },
      color: {
        field: 'status',
        type: 'nominal',
        scale: {
          domain: ['Critical', 'Off-track', 'Slipping', 'On-track'],
          range: ['#b91c1c', '#c2410c', '#a16207', '#166534']
        },
        legend: { title: 'Status' }
      },
      size: {
        field: 'slipDays',
        type: 'quantitative',
        scale: { range: [80, 400] },
        legend: { title: 'Slip days' }
      },
      tooltip: [
        { field: 'parentMilestone', title: 'Parent' },
        { field: 'fullMilestone', title: 'Milestone' },
        { field: 'status', title: 'Status' },
        { field: 'slipDays', title: 'Slip (days)' },
        { field: 'finish', type: 'temporal', title: 'Finish' }
      ]
    },
    config: {
      axis: { labelColor: '#334155', titleColor: '#334155' },
      view: { stroke: '#e2e8f0' },
      background: '#ffffff'
    }
  };

  vegaEmbed('#vegaMilestoneOverview', spec, { actions: false, renderer: 'svg' }).catch(() => {
    target.innerHTML = '<div style="font-size:12px;color:#64748b;padding:10px">Could not render overview chart.</div>';
  });
}

function renderVisMilestoneOverview() {
  const target = document.getElementById('visMilestoneOverview');
  const rows = getOverviewMilestoneRows();
  if (!target) return;
  if (!rows.length) {
    target.innerHTML = '<div style="font-size:12px;color:#64748b;padding:10px">No milestone data available.</div>';
    return;
  }
  if (typeof vis === 'undefined' || !vis.Timeline) {
    target.innerHTML = '<div style="font-size:12px;color:#64748b;padding:10px">vis-timeline failed to load.</div>';
    return;
  }

  const groupsOrder = [...new Set(rows.map((r) => r.parentMilestone))];
  const groups = new vis.DataSet(groupsOrder.map((g, idx) => ({ id: idx + 1, content: g, value: idx + 1 })));
  const groupByName = new Map(groupsOrder.map((g, idx) => [g, idx + 1]));

  const items = new vis.DataSet(rows.map((r) => ({
    id: r.id,
    uid: r.uid,
    group: groupByName.get(r.parentMilestone),
    start: r.finish,
    type: 'point',
    content: '<span style="font-size:11px;font-weight:700">' + esc(r.milestone) + '</span>',
    title: esc(r.fullMilestone + ' · ' + r.status + ' · ' + (r.slipDays > 0 ? '+' + r.slipDays + 'd' : '0d')),
    className: r.status === 'Critical' ? 'bar-critical' : (r.status === 'Off-track' ? 'bar-immediate' : (r.status === 'Slipping' ? 'bar-soon' : 'bar-watch'))
  })));

  const options = {
    stack: false,
    horizontalScroll: true,
    verticalScroll: true,
    zoomKey: 'ctrlKey',
    orientation: { axis: 'top' },
    margin: { item: 8, axis: 8 },
    height: '220px',
    showCurrentTime: true
  };

  target.innerHTML = '';
  visOverviewTimeline = new vis.Timeline(target, items, groups, options);
  visOverviewTimeline.on('select', (props) => {
    if (!props.items || !props.items.length) return;
    const selected = items.get(props.items[0]);
    if (selected && selected.uid) selectFocusMilestone(String(selected.uid));
  });
}

function renderMilestoneCommandCenter() {
  const rows = (m.decisionMetrics?.milestoneCommandTable || []).slice(0, 20);
  const el = document.getElementById('milestoneCommandCenter');
  if (!rows.length) {
    el.innerHTML = '<div class="cmd-wrap"><div class="cmd-head"><div class="cmd-title">Milestone Command Table</div><div class="cmd-sub">No milestone data available.</div></div></div>';
    return;
  }

  const statusClass = (status) => {
    const s = String(status || '').toLowerCase();
    if (s.includes('critical')) return 'critical';
    if (s.includes('off-track')) return 'offtrack';
    if (s.includes('slipping')) return 'slipping';
    return 'ontrack';
  };

  const tableRows = rows.map((r) => {
    const slipText = r.slipDays > 0 ? ('+' + r.slipDays + 'd') : '0d';
    const finishText = fmtD(r.finish);
    const clickAttr = ' onclick="selectFocusMilestone(' + Number(r.uid || 0) + ')" style="cursor:pointer" ';
    return '<tr' + clickAttr + '>'
      + '<td>' + esc(r.parentMilestone || 'Program') + '</td>'
      + '<td title="' + esc(r.name) + '">' + esc(shortText(r.name, 42)) + '</td>'
      + '<td><span class="cmd-badge ' + statusClass(r.status) + '">' + esc(r.status) + '</span></td>'
      + '<td>' + esc(slipText) + '</td>'
      + '<td>' + esc(finishText) + '</td>'
      + '<td class="cmd-impact">'+ esc(String(r.impact || 0)) + '</td>'
      + '</tr>';
  }).join('');

  el.innerHTML = '<div class="cmd-wrap">'
    + '<div class="cmd-head">'
    + '<div class="cmd-title">Milestone Command Table</div>'
    + '<div class="cmd-sub">Grouped by parent milestone (Text20 when available) · click row to focus timeline</div>'
    + '</div>'
    + '<div class="cmd-table-wrap">'
    + '<table class="cmd-table">'
    + '<thead><tr><th>Parent milestone</th><th>Milestone task</th><th>Status</th><th>Slip</th><th>Finish</th><th>Impact</th></tr></thead>'
    + '<tbody>' + tableRows + '</tbody>'
    + '</table>'
    + '</div>'
    + '</div>';
}

function renderDecisionStrip() {
  const d = m.decisionMetrics || {};
  const criticalChain = (d.criticalChain || []).map((t) => shortText(t.name, 28)).join(' → ');
  const dataLoadFocus = d.focusExamples?.loads || d.topDataLoadRisk;
  const testingFocus = d.focusExamples?.testing || d.topTestingRisk;
  const cards = [
    {
      eyebrow: 'Decision Now',
      headline: d.immediateExample ? shortText(d.immediateExample.name, 52) : 'No immediate decision called out',
      subline: d.immediateExample ? ('Finish ' + fmtD(d.immediateExample.finish) + (d.immediateExample.slipDays>0 ? (' · +' + d.immediateExample.slipDays + 'd slip') : '') + ' · ' + esc(d.immediateExample.resourceNames || 'No owner')) : 'No immediate task narrative available.',
      pill: d.immediateExample ? '<span class="pill pill-red">'+esc(d.immediateExample.tier || 'IMMEDIATE')+'</span>' : '',
      focusUid: d.immediateExample?.uid || ''
    },
    {
      eyebrow: 'Next Major Milestone',
      headline: d.nextMilestone ? shortText(d.nextMilestone.name, 52) : 'No upcoming milestone found',
      subline: d.nextMilestone ? ((d.nextMilestone.cycle || 'Program') + (d.nextMilestone.track ? ' · ' + d.nextMilestone.track : '') + ' · Finish ' + fmtD(d.nextMilestone.finish)) : 'Milestone data unavailable.',
      pill: d.nextMilestone ? '<span class="pill pill-amber">'+esc(d.nextMilestone.cycle || 'Milestone')+'</span>' : '',
      focusUid: d.nextMilestone?.uid || ''
    },
    {
      eyebrow: 'Data Loads Focus (ITC2 first)',
      headline: dataLoadFocus ? shortText(dataLoadFocus.name, 52) : 'No data-load milestone flagged',
      subline: dataLoadFocus ? ((dataLoadFocus.cycle || 'Program') + ' · Finish ' + fmtD(dataLoadFocus.finish) + (dataLoadFocus.slipDays>0 ? (' · +' + dataLoadFocus.slipDays + 'd slip') : '')) : 'Use this card as the anchor for ITC2 Data Loads.',
      pill: dataLoadFocus ? '<span class="pill pill-blue">'+esc((dataLoadFocus.cycle || 'Program') + ' Loads')+'</span>' : '',
      focusUid: dataLoadFocus?.uid || ''
    },
    {
      eyebrow: 'Testing / Driver Chain',
      headline: testingFocus ? shortText(testingFocus.name, 52) : (criticalChain || 'No critical chain found'),
      subline: testingFocus
        ? ((testingFocus.cycle || 'Program') + ' testing · Finish ' + fmtD(testingFocus.finish) + (testingFocus.slipDays>0 ? (' · +' + testingFocus.slipDays + 'd slip') : '') + (testingFocus.why ? (' · ' + esc(testingFocus.why)) : ''))
        : (criticalChain ? ('Critical chain: ' + esc(criticalChain)) : 'No current critical driver chain available.'),
      pill: testingFocus ? '<span class="pill pill-red">Testing</span>' : '<span class="pill pill-green">CPM chain</span>',
      focusUid: testingFocus?.uid || ''
    }
  ];

  document.getElementById('decisionStrip').innerHTML = cards.map((card) =>
    '<div class="decision-card' + (card.focusUid ? ' clickable' : '') + '"' + (card.focusUid ? ' data-focus-uid="' + card.focusUid + '"' : '') + '>'
      + '<div class="eyebrow">' + card.eyebrow + '</div>'
      + '<div class="headline">' + card.headline + '</div>'
      + '<div class="subline">' + card.subline + '</div>'
      + (card.pill || '')
    + '</div>'
  ).join('');
  document.querySelectorAll('#decisionStrip [data-focus-uid]').forEach((el) => {
    el.onclick = () => selectFocusMilestone(el.getAttribute('data-focus-uid'));
  });
}

function toggleSupplement(kind) {
  if (kind === 'decision') showDecisionStrip = !showDecisionStrip;
  if (kind === 'cpm') showCpmStrip = !showCpmStrip;
  applySupplementVisibility();
}

function toggleSimpleMode() {
  simpleMode = !simpleMode;
  applySupplementVisibility();
}

function applySupplementVisibility() {
  document.body.classList.toggle('simple-mode', simpleMode);
  document.getElementById('decisionStrip').style.display = showDecisionStrip ? 'grid' : 'none';
  document.getElementById('cpmExplain').style.display = showCpmStrip ? 'grid' : 'none';
  document.getElementById('criticalChainStrip').style.display = simpleMode ? 'none' : 'grid';
  document.getElementById('btnToggleDecision').className = showDecisionStrip ? 'active' : '';
  document.getElementById('btnToggleCpm').className = showCpmStrip ? 'active' : '';
  document.getElementById('btnSimpleMode').className = simpleMode ? 'active' : '';
  document.getElementById('btnSimpleMode').textContent = simpleMode ? 'Simple mode' : 'Analyst mode';
  document.getElementById('btnToggleTriage').className = 'layout-btn' + (triageVisible ? ' active' : '');
  document.getElementById('btnToggleTriage').textContent = triageVisible ? 'Hide triage' : 'Show triage';
  document.getElementById('layout').classList.toggle('triage-collapsed', simpleMode || !triageVisible);
}

function toggleTriagePanel() {
  triageVisible = !triageVisible;
  applySupplementVisibility();
}

function renderCriticalChainStrip() {
  const ccn = m.decisionMetrics?.criticalChainNarrative || null;
  const guide = m.decisionMetrics?.triageGuide || null;

  if (!ccn) {
    document.getElementById('criticalChainStrip').innerHTML = '';
    return;
  }

  // Build critical chain task sequence with at-risk highlighting
  const chainSeq = (ccn.chain || []).map((task) => {
    const isAtRisk = task.riskStatus && task.riskStatus !== 'ON_TRACK' && task.riskStatus !== 'MINOR_SLIP';
    return '<span class="chain-task' + (isAtRisk ? ' at-risk' : '') + '" title="' + esc(task.name + ' (' + (task.slipDays > 0 ? '+' + task.slipDays + 'd' : 'on track') + ')')
      + '">' + shortText(task.name, 20) + '</span>';
  }).join('<span class="chain-sep">→</span>');

  // Build chain stats
  const projectEnd = ccn.projectEndDate ? fmtD(ccn.projectEndDate) : 'Unknown';
  const cascadeCount = ccn.cascadeImpactIfFirstAtRiskSlips?.count || 0;
  const cascadeDirect = ccn.cascadeImpactIfFirstAtRiskSlips?.direct || 0;

  const chainPanel = '<div class="chain-panel">'
    + '<div class="title">🔴 YOUR CRITICAL PATH</div>'
    + '<div class="chain-sequence">' + chainSeq + '</div>'
    + '<div class="chain-stats">'
    + '<div class="chain-stat-row"><span>Tasks on chain:</span><span class="chain-stat-val">' + (ccn.chain.length || 0) + '</span></div>'
    + '<div class="chain-stat-row"><span>At risk:</span><span class="chain-stat-val">' + (ccn.atRiskCount || 0) + ' / ' + (ccn.chain.length || 0) + '</span></div>'
    + '<div class="chain-stat-row"><span>Project delivery:</span><span class="chain-stat-val">' + projectEnd + '</span></div>'
    + (cascadeCount > 0 ? '<div class="chain-stat-row"><span>If first at-risk slips:</span><span class="chain-stat-val">' + cascadeCount + ' tasks affected</span></div>' : '')
    + '</div>'
    + '</div>';

  // Build triage guide
  const triageItems = (guide?.guide || []).slice(0, 6).map((task) => {
    const isTier1 = task.triageTier === 'TIER1_CRITICAL';
    const isTier2 = task.triageTier === 'TIER2_BLOCKS_CRITICAL';
    const tierNum = isTier1 ? '1' : (isTier2 ? '2' : '3');
    const tierClass = isTier1 ? 'tier-1' : (isTier2 ? 'tier-2' : '');
    const tierIcon = isTier1 ? '🔴' : (isTier2 ? '🟡' : '🔵');
    return '<div class="triage-item ' + tierClass + '">'
      + '<span class="tier-icon">' + tierIcon + '</span>'
      + '<span class="triage-name" title="' + esc(task.name) + '">' + shortText(task.name, 28) + '</span>'
      + '<span class="triage-slip">' + (task.slipDays > 0 ? '+' + task.slipDays + 'd' : 'OK') + '</span>'
      + '</div>';
  }).join('');

  const triagePanel = '<div class="triage-panel">'
    + '<div class="title">📋 TRIAGE RANKING (Top 6)</div>'
    + '<div class="triage-list">'
    + '<div style="font-size:10px;color:#5c6f85;margin-bottom:4px">'
    + '<span style="color:#dc2626">🔴 Tier 1: ' + (guide?.tier1Count || 0) + '</span> · '
    + '<span style="color:#f97316">🟡 Tier 2: ' + (guide?.tier2Count || 0) + '</span> · '
    + '<span style="color:#1d4ed8">🔵 Tier 3: ' + (guide?.tier3Count || 0) + '</span>'
    + '</div>'
    + (triageItems || '<div style="color:#999;padding:6px">No IMMEDIATE tasks found (good news!)</div>')
    + '</div>'
    + '</div>';

  document.getElementById('criticalChainStrip').innerHTML = chainPanel + triagePanel;
}


function renderCpmExplain() {
  const chain = m.cpmBreakdown || [];
  const el = document.getElementById('cpmExplain');
  if (!chain.length) {
    el.innerHTML = '<div class="cpm-empty">No critical-path tasks found in current CPM run.</div>';
    return;
  }
  el.innerHTML = chain.map((task) => {
    const slip = task.slipDays > 0 ? (' · +' + task.slipDays + 'd slip') : '';
    const drivers = (task.drivenBy && task.drivenBy.length) ? task.drivenBy.slice(0, 2).join('; ') : 'No predecessor driver';
    const blocks = (task.blocks && task.blocks.length) ? task.blocks.slice(0, 2).join('; ') : 'No downstream task';
    return '<div class="cpm-card clickable" data-cpm-uid="' + task.uid + '">'
      + '<div class="eyebrow">CPM Chain Step ' + task.rank + ' · Slack ' + task.totalSlack + 'd</div>'
      + '<div class="title">' + esc(task.name) + '</div>'
      + '<div class="meta">Finish ' + fmtD(task.finish) + slip + ' · ' + esc(task.workstream || 'Unassigned') + '</div>'
      + '<div class="row"><span class="lbl">Why</span><span>' + esc(task.why || 'Zero slack critical-path task') + '</span></div>'
      + '<div class="row"><span class="lbl">Driver</span><span>' + esc(drivers) + '</span></div>'
      + '<div class="row"><span class="lbl">Impact</span><span>' + esc(blocks) + '</span></div>'
      + '<div class="row"><span class="lbl">Action</span><span>' + esc(task.action || 'Escalate and compress schedule') + '</span></div>'
      + '</div>';
  }).join('');
  document.querySelectorAll('#cpmExplain [data-cpm-uid]').forEach((card) => {
    card.onclick = () => focusTask(card.getAttribute('data-cpm-uid'), { driverOnly: true });
  });
}

function listTaskNames(ids) {
  if (!ids || !ids.length) return '—';
  return ids.map((id) => taskByUid.get(id)?.name).filter(Boolean).slice(0, 8).map((n) => '• ' + esc(n)).join('<br>') || '—';
}

// ── Action Panel ────────────────────────────────────────────────────────
let actionPanelOpen = false;
let actionPanelTaskUid = null;
let actionPanelLane = 'ALL';

function toggleActionPanel() {
  if (actionPanelOpen) closeActionPanel(); else openActionPanel();
}
function openActionPanel() {
  document.getElementById('actionPanel').classList.add('open');
  actionPanelOpen = true;
  renderActionPanel();
}
function closeActionPanel() {
  document.getElementById('actionPanel').classList.remove('open');
  actionPanelOpen = false;
}
function setApLane(lane) {
  actionPanelLane = lane;
  document.querySelectorAll('.ap-lane-chip').forEach(function(c) {
    c.classList.toggle('active', c.dataset.lane === lane);
  });
  renderActionPanel();
}

function renderActionPanel() {
  var items = (DATA.actionItems || []);
  var taskUidSet = null;

  // If a task is selected, filter to that task's items
  if (actionPanelTaskUid) {
    var selTask = taskByUid.get(String(actionPanelTaskUid));
    if (selTask) {
      taskUidSet = new Set([selTask.uid]);
    }
  } else {
    // Apply view filter: build set of qualifying task IDs
    var vf = document.getElementById('viewFilter') ? document.getElementById('viewFilter').value : 'ALL';
    if (vf === 'CRITICAL') {
      taskUidSet = new Set();
      DATA.tasks.forEach(function(t) { if (t.cpm && t.cpm.critical) taskUidSet.add(t.uid); });
    } else if (vf === 'DRIVER') {
      taskUidSet = new Set();
      DATA.tasks.forEach(function(t) { if (t.cpm && t.cpm.critical) taskUidSet.add(t.uid); });
    }
    // For other views: build set from currently visible gantt tasks if possible
  }

  // Filter by task set
  if (taskUidSet) {
    items = items.filter(function(a) { return a.taskId && taskUidSet.has(a.taskId); });
  }

  // Filter by lane
  if (actionPanelLane !== 'ALL') {
    var laneMap = { actNow: 'Act Now', atRisk: 'At Risk Soon', watchlist: 'Watchlist' };
    var laneTarget = laneMap[actionPanelLane];
    items = items.filter(function(a) { return a.attentionLane === laneTarget; });
  }

  // Sort: Act Now first, then At Risk Soon, then Watchlist; within each by severity
  var laneOrder = { 'Act Now': 0, 'At Risk Soon': 1, 'Watchlist': 2 };
  var sevOrder = function(s) {
    var sl = String(s || '').toLowerCase();
    if (/critical|very high|urgent/.test(sl)) return 0;
    if (/high/.test(sl)) return 1;
    if (/medium|med/.test(sl)) return 2;
    return 3;
  };
  items = items.slice().sort(function(a, b) {
    var la = laneOrder[a.attentionLane] || 2, lb = laneOrder[b.attentionLane] || 2;
    if (la !== lb) return la - lb;
    return sevOrder(a.severity) - sevOrder(b.severity);
  }).slice(0, 200);

  // Update header count + dot
  var count = items.length;
  var hc = document.getElementById('apHeaderCount');
  if (hc) hc.textContent = count;
  var dot = document.getElementById('apToggleDot');
  if (dot) dot.textContent = (DATA.actionItems || []).length;

  // Sub label
  var sub = document.getElementById('apSub');
  if (sub) {
    if (actionPanelTaskUid) {
      var st = taskByUid.get(String(actionPanelTaskUid));
      sub.textContent = 'Issues for: ' + (st ? (st.name || '').substring(0, 45) : '—');
    } else {
      var vf2 = document.getElementById('viewFilter') ? document.getElementById('viewFilter').value : 'ALL';
      var vfLabels = { ALL: 'all views', CRITICAL: 'Critical Path', DRIVER: 'Driver Flow', SLIPPING: 'Slipping tasks', PAST_DUE: 'Past due tasks', CRITICAL_TODAY: 'Needs action now', NEAR_CRITICAL: 'Near-critical tasks', SOON: 'Coming up soon', WATCH: 'Watch list' };
      sub.textContent = 'Showing issues for: ' + (vfLabels[vf2] || vf2);
    }
  }

  // Build cards
  var list = document.getElementById('apList');
  if (!list) return;
  if (!items.length) {
    list.innerHTML = '<div class="ap-empty">No issues found for current selection.<br><span style="font-size:10px;opacity:.6">Try changing the view filter or selecting a different task.</span></div>';
    return;
  }

  var fmtD = function(d) { return d ? new Date(d).toLocaleDateString('en-US', { month: 'short', day: 'numeric', year: '2-digit' }) : ''; };
  var laneClass = function(l) { return l === 'Act Now' ? 'lane-now' : l === 'At Risk Soon' ? 'lane-risk' : 'lane-watch'; };
  var laneBadge = function(l) { return l === 'Act Now' ? '<span class="ap-lane-badge now">Act Now</span>' : l === 'At Risk Soon' ? '<span class="ap-lane-badge risk">At Risk</span>' : '<span class="ap-lane-badge watch">Watch</span>'; };
  var sevClass = function(s) { var sl = String(s||'').toLowerCase(); return /critical|very high|urgent/.test(sl)?'critical':/high/.test(sl)?'high':/medium|med/.test(sl)?'medium':'low'; };

  list.innerHTML = items.map(function(a) {
    var linkedTask = a.taskId ? taskByUid.get(a.taskId) : null;
    var taskName = linkedTask ? (linkedTask.name || '').substring(0, 50) : '';
    var name = (a.name || a.title || 'Unnamed issue').substring(0, 100);
    var meta = [a.type, a.category].filter(Boolean).join(' · ');
    var mDate = a.milestoneDate ? 'Milestone: ' + fmtD(a.milestoneDate) : '';
    var lc = laneClass(a.attentionLane);
    var sc = sevClass(a.severity);
    return '<div class="ap-item ' + lc + '">'
      + '<div class="ap-item-top">' + laneBadge(a.attentionLane) + '<span class="ap-sev ' + sc + '">' + (a.severity || 'Low') + '</span></div>'
      + '<div class="ap-item-name" title="' + name.replace(/"/g,'&quot;') + '">' + name + '</div>'
      + (taskName ? '<div class="ap-item-task">&#128196; ' + taskName + '</div>' : '')
      + ((meta || mDate) ? '<div class="ap-item-meta">' + [meta, mDate].filter(Boolean).join(' &nbsp;·&nbsp; ') + '</div>' : '')
      + '</div>';
  }).join('');
}
// ── End Action Panel ─────────────────────────────────────────────────────
// ── Network Canvas ──// ── Network Canvas ──────────────────────────────────────────────────────
let networkCanvasOpen = false;
let networkCanvasMode = 'ALL';
let selectedCanvasMilestones = new Set();
let ncVisNetwork = null;

function openNetworkCanvas(uid) {
  const task = taskByUid.get(String(uid));
  if (!task) return;
  const vf = document.getElementById('viewFilter') ? document.getElementById('viewFilter').value : 'ALL';
  networkCanvasMode = vf === 'CRITICAL' ? 'CRITICAL' : vf === 'DRIVER' ? 'DRIVER' : 'ALL';
  selectedCanvasMilestones = new Set([String(uid)]);
  actionPanelTaskUid = String(uid);
  if (actionPanelOpen) renderActionPanel();
  document.getElementById('networkCanvas').classList.add('open');
  networkCanvasOpen = true;
  renderNetworkCanvas();
}

function closeNetworkCanvas() {
  document.getElementById('networkCanvas').classList.remove('open');
  networkCanvasOpen = false;
  selectedCanvasMilestones.clear();
  if (ncVisNetwork) { ncVisNetwork.destroy(); ncVisNetwork = null; }
}

function renderNetworkCanvas() {
  renderNcMilestoneStrip();
  const nd = buildNetworkData();
  renderNcGraph(nd.nodes, nd.edges);
  renderNcCallouts();
  const modeLabels = { ALL: 'All milestones', CRITICAL: 'Critical path', DRIVER: 'Driver flow' };
  const chip = document.getElementById('ncModeChip');
  if (chip) chip.textContent = 'Canvas: ' + (modeLabels[networkCanvasMode] || networkCanvasMode);
}

function ncFmtDate(d) {
  if (!d) return '—';
  return new Date(d).toLocaleDateString('en-US', { month: 'short', day: 'numeric' });
}

function ncSlipDays(t) {
  if (!t.finish || !t.baselineFinish) return null;
  const diff = Math.round((new Date(t.finish) - new Date(t.baselineFinish)) / 86400000);
  return diff > 0 ? diff : null;
}

function renderNcMilestoneStrip() {
  const strip = document.getElementById('ncMilestoneStrip');
  if (!strip) return;
  let html = '';
  for (const uid of selectedCanvasMilestones) {
    const t = taskByUid.get(uid);
    if (!t) continue;
    const isCp = t.cpm && t.cpm.critical;
    const isDriver = t.cpmDriverScore > 0;
    const slip = ncSlipDays(t);
    const cardCls = isCp ? 'nc-mcard cp-card' : isDriver ? 'nc-mcard driver-card' : 'nc-mcard';
    const badgeHtml = (isCp ? '<span class="cmd-badge critical" style="font-size:9px">CP</span>' : '')
      + (isDriver ? '<span class="cmd-badge offtrack" style="font-size:9px">Driver</span>' : '')
      + (t.milestone ? '<span class="cmd-badge ontrack" style="font-size:9px">Milestone</span>' : '')
      + (slip ? '<span class="cmd-badge slipping" style="font-size:9px">+' + slip + 'd</span>' : '');
    const name = t.name || '';
    const nameShort = name.length > 28 ? name.substring(0, 28) + '\u2026' : name;
    const nameEsc = name.replace(/"/g, '&quot;');
    html += '<div class="' + cardCls + '" title="' + nameEsc + '">';
    html += '<div class="nc-mcard-top"><span class="nc-mcard-ws">' + (t.workstream || '') + '</span>';
    html += '<span class="nc-mcard-sf">' + ncFmtDate(t.start) + ' \u2192 ' + ncFmtDate(t.finish) + '</span></div>';
    if (badgeHtml) html += '<div class="nc-mcard-badges">' + badgeHtml + '</div>';
    html += '<div class="nc-mcard-name" title="' + nameEsc + '">' + nameShort + '</div></div>';
  }
  strip.innerHTML = html || '<span style="color:#94a3b8;font-size:11px;padding:4px 0">Click a bar on the Gantt to open canvas</span>';
}

function buildNetworkData() {
  const nodes = [];
  const edges = [];
  const seen = new Set();
  const MAX_DEPTH = 4;

  function shouldInclude(t) {
    if (!t) return false;
    if (networkCanvasMode === 'CRITICAL') return !!(t.cpm && t.cpm.critical);
    if (networkCanvasMode === 'DRIVER') return !!(t.cpmDriverScore > 0 || (t.cpm && t.cpm.critical));
    return true;
  }

  function nodeColor(t) {
    const isCp = t.cpm && t.cpm.critical;
    const isDriver = t.cpmDriverScore > 0;
    const slip = ncSlipDays(t);
    const bg = slip ? '#fef3c7' : isCp ? '#fef2f2' : isDriver ? '#eff6ff' : '#f8fafc';
    const border = isCp ? '#dc2626' : isDriver ? '#1d4ed8' : slip ? '#d97706' : '#94a3b8';
    return { background: bg, border: border, highlight: { background: bg, border: '#0f172a' } };
  }

  function addNode(t) {
    const key = String(t.uid);
    if (seen.has(key)) return;
    seen.add(key);
    const name = t.name || '';
    const label = name.length > 22 ? name.substring(0, 22) + '\u2026' : name;
    const slip = ncSlipDays(t);
    const titleText = name + '\\n' + (t.workstream || '') + '\\nS: ' + ncFmtDate(t.start) + '  F: ' + ncFmtDate(t.finish)
      + (slip ? '\\nSlip: +' + slip + 'd' : '') + (t.cpm ? '\\nSlack: ' + (t.cpm.totalSlack != null ? t.cpm.totalSlack : '—') + 'd' : '');
    nodes.push({
      id: key, label: label, title: titleText,
      shape: t.milestone ? 'diamond' : 'box',
      size: t.milestone ? 22 : 14,
      color: nodeColor(t),
      font: { size: t.milestone ? 12 : 11 },
      borderWidth: selectedCanvasMilestones.has(key) ? 3 : 1
    });
  }

  function bfs(startUid, direction, depth) {
    if (depth <= 0) return;
    const t = taskByUid.get(String(startUid));
    if (!t) return;
    const rawLinks = direction === 'up' ? (t.predecessors || t.predIds || []) : (t.successors || t.sucIds || []);
    const links = rawLinks.map(function(x) { return (x && typeof x === 'object') ? (x.uid || x.id || x) : x; });
    for (let i = 0; i < links.length; i++) {
      const predUid = String(links[i]);
      const pt = taskByUid.get(predUid);
      if (!pt || !shouldInclude(pt)) continue;
      const ptKey = String(pt.uid);
      if (!seen.has(ptKey)) {
        addNode(pt);
        if (direction === 'up') {
          edges.push({ from: ptKey, to: String(startUid), arrows: 'to', color: { color: '#94a3b8' } });
        } else {
          edges.push({ from: String(startUid), to: ptKey, arrows: 'to', color: { color: '#94a3b8' } });
        }
      }
      bfs(predUid, direction, depth - 1);
    }
  }

  for (const uid of selectedCanvasMilestones) {
    const t = taskByUid.get(uid);
    if (!t) continue;
    addNode(t);
    bfs(uid, 'up', MAX_DEPTH);
    bfs(uid, 'down', MAX_DEPTH);
  }
  return { nodes: nodes, edges: edges };
}

function renderNcGraph(nodes, edges) {
  const container = document.getElementById('ncGraph');
  if (!container || !window.vis || typeof vis.Network === 'undefined') {
    if (container) container.innerHTML = '<div style="padding:16px;color:#94a3b8;font-size:12px">vis-network not loaded — check CDN connection</div>';
    return;
  }
  const data = { nodes: new vis.DataSet(nodes), edges: new vis.DataSet(edges) };
  const options = {
    layout: {
      hierarchical: {
        enabled: true, direction: 'LR', sortMethod: 'directed',
        levelSeparation: 180, nodeSpacing: 80, treeSpacing: 120,
        blockShifting: true, edgeMinimization: true, parentCentralization: true
      }
    },
    physics: { enabled: false },
    interaction: { hover: true, tooltipDelay: 200, dragView: true, zoomView: true },
    edges: { arrows: { to: { enabled: true, scaleFactor: 0.7 } }, smooth: { type: 'cubicBezier' }, width: 1.5 },
    nodes: { font: { size: 11, face: 'Segoe UI,system-ui,sans-serif' }, widthConstraint: { maximum: 160 } }
  };
  if (ncVisNetwork) { ncVisNetwork.destroy(); ncVisNetwork = null; }
  ncVisNetwork = new vis.Network(container, data, options);
  ncVisNetwork.once('afterDrawing', function() { ncVisNetwork.fit({ animation: false }); });
  ncVisNetwork.on('click', function(params) {
    if (params.nodes && params.nodes.length) {
      selectedCanvasMilestones.add(String(params.nodes[0]));
      renderNetworkCanvas();
    }
  });
}

function renderNcCallouts() {
  const el = document.getElementById('ncCallouts');
  if (!el) return;
  const uids = Array.from(selectedCanvasMilestones);
  if (!uids.length) { el.innerHTML = ''; return; }
  const t = taskByUid.get(uids[0]);
  if (!t) { el.innerHTML = ''; return; }
  const rawPreds = t.predecessors || t.predIds || [];
  const rawSuccs = t.successors || t.sucIds || [];
  const preds = rawPreds.map(function(u) { return taskByUid.get(String((u && typeof u === 'object') ? (u.uid || u.id || u) : u)); }).filter(Boolean);
  const succs = rawSuccs.map(function(u) { return taskByUid.get(String((u && typeof u === 'object') ? (u.uid || u.id || u) : u)); }).filter(Boolean);
  const slip = ncSlipDays(t);
  const driversText = preds.length ? preds.slice(0, 3).map(function(p) { return (p.name || '').substring(0, 40); }).join('; ') : 'No upstream predecessors found';
  const impactText = succs.length ? succs.slice(0, 3).map(function(s) { return (s.name || '').substring(0, 40); }).join('; ') : 'No downstream successors found';
  const consequenceText = slip
    ? 'This task is running +' + slip + ' days late. Downstream tasks are at risk of cascading delay.'
    : (t.cpm && t.cpm.critical) ? 'On the critical path — any further delay extends the project end date directly.'
    : 'Currently within slack tolerance but monitor closely.';
  const nextText = slip ? 'Escalate immediately. Confirm recovery plan and revised finish date.'
    : 'Review predecessor progress and confirm no emerging constraints.';
  el.innerHTML = '<div class="nc-callout"><strong>What drives this</strong>' + driversText + '</div>'
    + '<div class="nc-callout"><strong>What this affects</strong>' + impactText + '</div>'
    + '<div class="nc-callout"><strong>If not resolved</strong>' + consequenceText + '</div>'
    + '<div class="nc-callout"><strong>Next action</strong>' + nextText + '</div>';
}
// ── End Network Canvas ───────────────────────────────────────────────────
function openNetworkModal(uid) {
  const task = taskByUid.get(uid);
  if (!task) return;
  window.__networkUid = uid;
  const narrative = taskNarrativeMap.get(uid);
  document.getElementById('networkTitle').textContent = task.name;
  document.getElementById('networkSubtitle').textContent = (task.outlineNumber || 'WBS ?') + ' · ' + (task.workstream || 'Unassigned') + ' · Slack ' + (task.cpm?.totalSlack ?? '—') + 'd';
  document.getElementById('networkSummary').innerHTML =
    'Start: ' + esc(task.start || '—') + '<br>'
    + 'Finish: ' + esc(task.finish || '—') + '<br>'
    + 'Baseline finish: ' + esc(task.baselineFinish || '—') + '<br>'
    + 'Owner: ' + esc(task.resourceNames || 'Unassigned');
  document.getElementById('networkDrivers').innerHTML = listTaskNames(task.predecessors || []);
  document.getElementById('networkImpacts').innerHTML = listTaskNames(task.successors || []);
  document.getElementById('networkNarrative').innerHTML =
    '<b>What:</b> ' + esc(narrative?.whatIsWrong || (task.cpm?.critical ? 'On critical path' : '')) + '<br>'
    + '<b>Why:</b> ' + esc(narrative?.why || '') + '<br>'
    + '<b>Impact:</b> ' + esc(narrative?.impact || '') + '<br>'
    + '<b>Action:</b> ' + esc(narrative?.resolution || 'Monitor and recover');
  document.getElementById('networkModal').classList.add('open');
}

function closeNetworkModal() {
  document.getElementById('networkModal').classList.remove('open');
}

function getVisibleFocusMilestones() {
  const visibleAnchors = laneAnchors
    .filter((anchor) => selectedCycles.has(anchor.cycle) && selectedTracks.has(anchor.track));
  const candidates = [];
  for (const anchor of visibleAnchors) {
    const prefix = String(anchor.outlineNumber || '') + '.';
    for (const task of DATA.tasks) {
      const outline = String(task.outlineNumber || '');
      if (!(task.uid === anchor.uid || outline.startsWith(prefix))) continue;
      if (!task.milestone) continue;
      if (!task.start || !task.finish) continue;
      candidates.push({
        uid: String(task.uid),
        label: anchor.cycle + ' · ' + anchor.track + ' — ' + shortText(task.name, 54),
        cycle: anchor.cycle,
        track: anchor.track,
        outlineNumber: task.outlineNumber,
        finish: task.finish
      });
    }
  }
  const dedup = new Map();
  for (const item of candidates.sort((a, b) => cycleOrder.indexOf(a.cycle) - cycleOrder.indexOf(b.cycle) || trackOrder.indexOf(a.track) - trackOrder.indexOf(b.track) || String(a.outlineNumber).localeCompare(String(b.outlineNumber)))) {
    if (!dedup.has(item.uid)) dedup.set(item.uid, item);
  }
  return [...dedup.values()];
}

function renderFocusOptions() {
  const select = document.getElementById('focusMilestone');
  const options = getVisibleFocusMilestones();
  if (focusedMilestoneUid && !options.some((x) => x.uid === focusedMilestoneUid)) focusedMilestoneUid = '';
  select.innerHTML = '<option value="">All visible lanes</option>' + options.map((x) => '<option value="' + x.uid + '"' + (x.uid === focusedMilestoneUid ? ' selected' : '') + '>' + esc(x.label) + '</option>').join('');
}

function selectFocusMilestone(uid) {
  focusedMilestoneUid = uid || '';
  focusedTaskUid = focusedMilestoneUid;
  renderFocusOptions();
  rebuildGantt();
}

function focusTask(uid, options = {}) {
  focusedTaskUid = uid || '';
  focusedMilestoneUid = uid || '';
  if (options.driverOnly === true) focusDriversOnly = true;
  if (options.driverOnly === false) focusDriversOnly = false;
  renderLaneChips();
  renderFocusOptions();
  rebuildGantt();
}

function renderLaneChips() {
  const cycleChips = document.getElementById('cycleChips');
  const trackChips = document.getElementById('trackChips');
  const btnContext = document.getElementById('btnContext');
  const btnFocusDrivers = document.getElementById('btnFocusDrivers');
  
  if (cycleChips) {
    cycleChips.innerHTML = cycleOrder.map((cycle) =>
      '<button class="filter-chip ' + (selectedCycles.has(cycle) ? 'active' : '') + '" data-cycle="' + cycle + '">' + cycle + '</button>'
    ).join('');
    document.querySelectorAll('#cycleChips [data-cycle]').forEach((btn) => {
      btn.onclick = () => toggleCycle(btn.getAttribute('data-cycle'));
    });
  }
  
  if (trackChips) {
    trackChips.innerHTML = trackOrder.map((track) =>
      '<button class="filter-chip track ' + (selectedTracks.has(track) ? 'active' : '') + '" data-track="' + track + '">' + track + '</button>'
    ).join('');
    document.querySelectorAll('#trackChips [data-track]').forEach((btn) => {
      btn.onclick = () => toggleTrack(btn.getAttribute('data-track'));
    });
  }
  
  if (btnContext) btnContext.className = 'filter-chip toggle' + (includeContext ? ' active' : '');
  if (btnFocusDrivers) btnFocusDrivers.className = 'filter-chip toggle' + (focusDriversOnly ? ' active' : '');
  document.getElementById('btnCriticalNetwork').className = 'filter-chip toggle' + (criticalNetworkOnly ? ' active' : '');
}

function toggleCycle(cycle) {
  if (selectedCycles.has(cycle) && selectedCycles.size > 1) selectedCycles.delete(cycle);
  else selectedCycles.add(cycle);
  renderLaneChips();
  renderFocusOptions();
  rebuildGantt();
}

function toggleTrack(track) {
  if (selectedTracks.has(track) && selectedTracks.size > 1) selectedTracks.delete(track);
  else selectedTracks.add(track);
  renderLaneChips();
  renderFocusOptions();
  rebuildGantt();
}

function toggleContextMode() {
  includeContext = !includeContext;
  renderLaneChips();
  refreshControlTooltips();
  rebuildGantt();
}

function toggleFocusDriversOnly() {
  focusDriversOnly = !focusDriversOnly;
  renderLaneChips();
  refreshControlTooltips();
  rebuildGantt();
}

function toggleCriticalNetwork() {
  criticalNetworkOnly = !criticalNetworkOnly;
  renderLaneChips();
  refreshControlTooltips();
  rebuildGantt();
}

function switchTier(tier) {
  activeTier = tier;
  ['IMMEDIATE','SOON','WATCH'].forEach(t => {
    const btn = document.getElementById('tab'+t[0]+t.slice(1).toLowerCase());
    if (btn) btn.className = 'tab-btn' + (t===tier ? (' active-'+(t==='IMMEDIATE'?'red':t==='SOON'?'amber':'blue')) : '');
  });
  document.getElementById('tabImmediate').className = 'tab-btn' + (tier==='IMMEDIATE' ? ' active-red' : '');
  document.getElementById('tabSoon').className = 'tab-btn' + (tier==='SOON' ? ' active-amber' : '');
  document.getElementById('tabWatch').className = 'tab-btn' + (tier==='WATCH' ? ' active-blue' : '');
  renderCards();
}

function renderCards() {
  const triageCardsEl = document.getElementById('triageCards');
  if (!triageCardsEl) return; // Element doesn't exist in new layout
  const key = activeTier === 'IMMEDIATE' ? 'immediate' : activeTier === 'SOON' ? 'soon' : 'watch';
  const tasks = m[key] || [];
  if (!tasks.length) {
    triageCardsEl.innerHTML = '<div style="padding:24px;text-align:center;color:var(--muted)">No tasks in this tier.</div>';
    return;
  }
  triageCardsEl.innerHTML = tasks.map((t, i) => \`
    <div class="task-card tier-\${t.tier}" id="card-\${t.uid}" onclick="toggleCard(\${t.uid}, this)" data-uid="\${t.uid}">
      <div class="card-header">
        <div>
          <div class="card-wbs">WBS \${esc(t.wbs||t.outlineNumber)} · \${esc(t.workstream||'Unassigned')}</div>
          <div class="card-name">\${esc(t.name)}</div>
        </div>
        <span class="tier-badge">\${t.tier}\${t.cpm&&t.cpm.critical?' · CRITICAL':''}</span>
      </div>
      <div style="font-size:11px;color:var(--muted)">
        Finish: \${fmtD(t.finish)}\${t.slipDays>0?' · <b style="color:var(--red)">+'+t.slipDays+' d slip</b>':''} · \${esc(t.resourceNames||'No owner')}
      </div>
      <div class="narrative">
        <div style="height:1px;background:var(--line);margin:8px 0"></div>
        <div class="nar-row"><span class="nar-label">What</span><span class="nar-val">\${esc(t.whatIsWrong)}</span></div>
        <div class="nar-row"><span class="nar-label">Why</span><span class="nar-val">\${esc(t.why)}</span></div>
        <div class="nar-row"><span class="nar-label">Impact</span><span class="nar-val">\${esc(t.impact)}</span></div>
        <div class="nar-row"><span class="nar-label">Resolution</span><span class="nar-val">\${esc(t.resolution)}</span></div>
      </div>
    </div>
  \`).join('');
}

function toggleCard(uid, el) {
  el.classList.toggle('expanded');
}

renderDecisionStrip();
renderCpmExplain();
renderCriticalChainStrip();
renderMilestoneCommandCenter();
renderVegaMilestoneOverview();
renderVisMilestoneOverview();
renderOverviewTabs();
setOverviewRenderer('vega');
renderCards();
renderLaneChips();
renderFocusOptions();
applySupplementVisibility();

// ── Frappe Gantt ────────────────────────────────────────────────────────────
let ganttChart = null;

// Populate workstream filter
const allWs = [...new Set(DATA.tasks.filter(t => t.workstream).map(t => t.workstream))].sort();
const wsSelect = document.getElementById('filterWs');
allWs.forEach(ws => { const o = document.createElement('option'); o.value = ws; o.text = ws; wsSelect.appendChild(o); });

function compareOutline(a, b) {
  return (a.outlineNumber || '').localeCompare(b.outlineNumber || '') || a.uid - b.uid;
}

function inferCycleClient(task) {
  const outline = String(task.outlineNumber || '');
  const name = String(task.name || '');
  if (outline.startsWith('1.2.1') || outline.startsWith('1.3.3.5') || outline.startsWith('1.1.1.4') || /\bITC[- ]?1\b/i.test(name)) return 'ITC1';
  if (outline.startsWith('1.2.2') || outline.startsWith('1.3.3.6') || outline.startsWith('1.1.1.5') || /\bITC[- ]?2\b/i.test(name)) return 'ITC2';
  if (outline.startsWith('1.2.3') || outline.startsWith('1.3.3.7') || outline.startsWith('1.1.1.6') || /\bITC[- ]?3\b/i.test(name)) return 'ITC3';
  if (outline.startsWith('1.2.4') || outline.startsWith('1.3.3.8') || outline.startsWith('1.1.1.7') || /\bUAT\b/i.test(name)) return 'UAT';
  return null;
}

function toIsoEnd(start, finish) {
  if (!start || !finish) return finish || start;
  if (finish > start) return finish;
  const d = new Date(start);
  d.setDate(d.getDate() + 1);
  return d.toISOString().slice(0, 10);
}

function dateBounds(rows) {
  const dated = rows.filter((t) => t.start && t.finish);
  if (!dated.length) {
    const today = new Date().toISOString().slice(0, 10);
    return { start: today, end: today };
  }
  const starts = dated.map((t) => new Date(t.start).getTime()).filter(Number.isFinite);
  const ends = dated.map((t) => new Date(toIsoEnd(t.start, t.finish)).getTime()).filter(Number.isFinite);
  return {
    start: new Date(Math.min(...starts)).toISOString().slice(0, 10),
    end: new Date(Math.max(...ends)).toISOString().slice(0, 10)
  };
}

function tierCls(uid, isSummary, isMilestone, isCritical) {
  if (isCritical) return 'bar-critical';
  if (isSummary)  return 'bar-track';
  const tier = tierMap.get(uid);
  if (tier === 'IMMEDIATE') return 'bar-immediate';
  if (tier === 'SOON')      return 'bar-soon';
  if (tier === 'WATCH')     return 'bar-watch';
  return 'bar-summary';
}

// MSP-style: complete=gray, critical=red, late=red, in-progress=dark blue, not-started=blue
function mspCls(task, isSummary, isCritical) {
  if (isCritical) return 'bar-critical';
  if (isSummary)  return 'bar-track';
  const pct    = task.percentComplete || 0;
  const today  = new Date().toISOString().slice(0, 10);
  const finish = task.finish ? String(task.finish).slice(0, 10) : '';
  if (pct >= 100)               return 'bar-msp-complete';
  if (finish && finish < today) return 'bar-msp-late';
  if (pct > 0)                  return 'bar-msp-inprogress';
  return 'bar-msp-notstarted';
}

// ISO week number (1-53)
function getWeekNum(d) {
  const jan1 = new Date(d.getFullYear(), 0, 1);
  return Math.ceil((((d - jan1) / 86400000) + jan1.getDay() + 1) / 7);
}

// Build view_modes array with custom Day/Week headers (Wk N / Month Year).
// The selected mode goes first so Frappe uses it as the initial view.
function buildViewModes(selectedMode) {
  const MONTHS = ['Jan','Feb','Mar','Apr','May','Jun','Jul','Aug','Sep','Oct','Nov','Dec'];
  const customDay = {
    name: 'Day', padding: '7d', step: '1d', date_format: 'YYYY-MM-DD',
    column_width: 30,
    lower_text: (d) => String(d.getDate()),
    upper_text: (d, ld) => (d.getDay() === 1 || !ld) ? 'Wk\u00a0' + getWeekNum(d) : '',
    upper_text_frequency: 7,
  };
  const customWeek = {
    name: 'Week', padding: '1m', step: '1d', date_format: 'YYYY-MM-DD',
    column_width: 38,
    lower_text: (d) => d.getDay() === 1 ? 'Wk\u00a0' + getWeekNum(d) : '',
    upper_text: (d, ld) => !ld || d.getMonth() !== ld.getMonth() ? MONTHS[d.getMonth()] + ' ' + d.getFullYear() : '',
    upper_text_frequency: 4,
  };
  const modeMap = { 'Day': customDay, 'Week': customWeek };
  const order = ['Day','Week','Month','Quarter Day','Half Day','Year'];
  const sorted = [selectedMode, ...order.filter(m => m !== selectedMode)];
  return sorted.map(m => modeMap[m] || m);
}

// Inject transparent colored bands per cycle behind gantt bars.
let lastGanttTasks = [];
function paintSwimlaneBands(tasks) {
  if (tasks) lastGanttTasks = tasks;
  setTimeout(() => {
    const svg = document.querySelector('#gantt');
    if (!svg) return;
    svg.querySelectorAll('.sl-band').forEach(e => e.remove());

    const idCycleMap = new Map();
    lastGanttTasks.forEach(t => {
      if (t._meta && !t._meta.synthetic && t._meta.cycle) {
        idCycleMap.set(String(t.id), t._meta.cycle);
      }
    });
    if (!idCycleMap.size) return;

    const cycleYBounds = new Map();
    svg.querySelectorAll('.bar-wrapper').forEach(wrapper => {
      const tid   = wrapper.getAttribute('data-id');
      const cycle = idCycleMap.get(tid);
      if (!cycle) return;
      const bar = wrapper.querySelector('.bar');
      if (!bar) return;
      const y = parseFloat(bar.getAttribute('y') || 0);
      const h = parseFloat(bar.getAttribute('height') || 20);
      if (!cycleYBounds.has(cycle)) cycleYBounds.set(cycle, { min: y, max: y + h });
      else {
        const b = cycleYBounds.get(cycle);
        b.min = Math.min(b.min, y);
        b.max = Math.max(b.max, y + h);
      }
    });
    if (!cycleYBounds.size) return;

    // Determine total SVG width from the furthest bar edge
    let svgW = 4000;
    svg.querySelectorAll('.bar').forEach(bar => {
      const rx = parseFloat(bar.getAttribute('x') || 0) + parseFloat(bar.getAttribute('width') || 0);
      if (rx > svgW) svgW = rx + 300;
    });

    const FILL   = { ITC1:'rgba(59,130,246,0.07)',  ITC2:'rgba(16,185,129,0.07)',  ITC3:'rgba(245,158,11,0.07)',  UAT:'rgba(139,92,246,0.07)' };
    const STROKE = { ITC1:'rgba(59,130,246,0.28)',  ITC2:'rgba(16,185,129,0.28)',  ITC3:'rgba(245,158,11,0.28)',  UAT:'rgba(139,92,246,0.28)' };
    const PAD = 7;

    const entries = [...cycleYBounds.entries()].reverse();
    for (const [cycle, bounds] of entries) {
      const rect = document.createElementNS('http://www.w3.org/2000/svg', 'rect');
      rect.setAttribute('class', 'sl-band');
      rect.setAttribute('x', '0');
      rect.setAttribute('y', String(bounds.min - PAD));
      rect.setAttribute('width', String(svgW));
      rect.setAttribute('height', String(bounds.max - bounds.min + PAD * 2));
      rect.setAttribute('fill', FILL[cycle] || 'rgba(100,100,100,0.05)');
      rect.setAttribute('stroke', STROKE[cycle] || 'rgba(100,100,100,0.15)');
      rect.setAttribute('stroke-width', '0.5');
      rect.setAttribute('rx', '4');
      rect.setAttribute('pointer-events', 'none');
      svg.insertBefore(rect, svg.firstChild);
    }
  }, 80);
}

function toGanttTask(task, customClass, metaExtras = {}) {
  return {
    id: String(task.uid),
    name: task.name,
    start: task.start,
    end: toIsoEnd(task.start, task.finish),
    progress: Math.min(100, Math.max(0, task.percentComplete || 0)),
    dependencies: '',
    custom_class: customClass || mspCls(task, task.summary, !!(task.cpm && task.cpm.critical)),
    _meta: {
      uid: task.uid,
      wbs: task.wbs || task.outlineNumber,
      workstream: task.workstream,
      resource: task.resourceNames,
      baselineFinish: task.baselineFinish,
      slipDays: taskNarrativeMap.get(task.uid)?.slipDays || 0,
      tier: tierMap.get(task.uid),
      whatIsWrong: taskNarrativeMap.get(task.uid)?.whatIsWrong || '',
      why: taskNarrativeMap.get(task.uid)?.why || '',
      impact: taskNarrativeMap.get(task.uid)?.impact || '',
      resolution: taskNarrativeMap.get(task.uid)?.resolution || '',
      critical: !!(task.cpm && task.cpm.critical),
      predecessorIds: task.predecessors || [],
      successorIds: task.successors || [],
      cycle: metaExtras.cycle || null,
      track: metaExtras.track || null,
      relation: metaExtras.relation || null,
      synthetic: false
    }
  };
}

function makeSyntheticRow(id, label, customClass, rows, meta = {}) {
  const bounds = dateBounds(rows);
  return {
    id,
    name: label,
    start: bounds.start,
    end: toIsoEnd(bounds.start, bounds.end),
    progress: 0,
    dependencies: '',
    custom_class: customClass,
    _meta: { synthetic: true, ...meta }
  };
}

function buildAllMilestoneRows() {
  const milestones = DATA.tasks
    .filter((t) => t.milestone && t.active && t.start && t.finish)
    .sort((a, b) => String(a.finish).localeCompare(String(b.finish)) || compareOutline(a, b));
  if (!milestones.length) return [];
  const rows = [makeSyntheticRow('MS-ALL', 'All milestones', 'bar-lane', milestones, { lane: 'Milestones' })];
  milestones.forEach((task) => rows.push(toGanttTask(task, 'bar-focus', { relation: 'milestone' })));
  return rows;
}

function buildCriticalPathRows() {
  const critical = DATA.tasks
    .filter((t) => t.cpm?.critical && !t.summary && t.active && t.start && t.finish)
    .sort((a, b) => (a.cpm?.earlyStart ?? 0) - (b.cpm?.earlyStart ?? 0) || compareOutline(a, b));
  if (!critical.length) return [];
  const rows = [makeSyntheticRow('CP-ROW', 'Critical path', 'bar-lane', critical, { lane: 'Critical Path' })];
  critical.forEach((task) => rows.push(toGanttTask(task, 'bar-critical', { relation: 'critical' })));
  return rows;
}

function buildDriverFlowRows() {
  const byUid = new Map(DATA.tasks.map((t) => [t.uid, t]));
  const critical = DATA.tasks.filter((t) => t.cpm?.critical && !t.summary && t.active);
  const keep = new Set(critical.map((t) => t.uid));
  const queue = [...keep];
  while (queue.length) {
    const uid = queue.shift();
    const t = byUid.get(uid);
    if (!t) continue;
    (t.predecessors || []).forEach((p) => {
      if (keep.has(p)) return;
      const pred = byUid.get(p);
      if (!pred || pred.summary || !pred.active) return;
      keep.add(p);
      queue.push(p);
    });
  }
  const driverRows = [...keep]
    .map((uid) => byUid.get(uid))
    .filter((t) => t && t.start && t.finish)
    .sort((a, b) => (a.cpm?.earlyStart ?? 0) - (b.cpm?.earlyStart ?? 0) || compareOutline(a, b));
  if (!driverRows.length) return [];
  const rows = [makeSyntheticRow('DRV-ROW', 'Driver flow (critical + predecessors)', 'bar-lane', driverRows, { lane: 'Driver Flow' })];
  driverRows.forEach((task) => {
    const cls = task.cpm?.critical ? 'bar-critical' : 'context-driver';
    rows.push(toGanttTask(task, cls, { relation: task.cpm?.critical ? 'critical' : 'driver' }));
  });
  return rows;
}

function slipDaysForTask(task) {
  const fromNarrative = taskNarrativeMap.get(task.uid)?.slipDays;
  if (Number.isFinite(fromNarrative)) return fromNarrative;
  if (!task.finish || !task.baselineFinish) return 0;
  const slip = Math.round((new Date(task.finish) - new Date(task.baselineFinish)) / 86400000);
  return slip > 0 ? slip : 0;
}

function modeLabel(mode) {
  const labels = {
    ALL: 'All key milestones',
    CRITICAL: 'Critical path',
    DRIVER: 'Driver flow',
    SLIPPING: 'Running late (slipping)',
    PAST_DUE: 'Past due',
    CRITICAL_TODAY: 'Needs action now',
    NEAR_CRITICAL: 'Needs attention soon',
    SOON: 'Coming up soon',
    WATCH: 'Watch list'
  };
  return labels[mode] || mode;
}

function buildModeFilteredRows(mode, filterWs, filterTierVal) {
  const today = new Date().toISOString().slice(0, 10);
  const tasks = DATA.tasks
    .filter((t) => !t.summary && t.active && t.start && t.finish)
    .filter((t) => passesFilters(t, filterWs, filterTierVal))
    .filter((t) => {
      if (mode === 'SLIPPING') return slipDaysForTask(t) > 0;
      if (mode === 'PAST_DUE') return (t.percentComplete || 0) < 100 && String(t.finish).slice(0, 10) < today;
      if (mode === 'CRITICAL_TODAY') {
        const tier = tierMap.get(t.uid);
        return !!t.cpm?.critical || tier === 'IMMEDIATE';
      }
      if (mode === 'NEAR_CRITICAL') {
        const slack = Number(t.cpm?.totalSlack);
        return Number.isFinite(slack) && slack > 0 && slack <= 10;
      }
      if (mode === 'SOON') return tierMap.get(t.uid) === 'SOON';
      if (mode === 'WATCH') return tierMap.get(t.uid) === 'WATCH';
      return false;
    })
    .sort((a, b) => String(a.finish).localeCompare(String(b.finish)) || compareOutline(a, b));

  if (!tasks.length) return [];

  const cycleIndex = new Map(cycleOrder.map((c, i) => [c, i]));
  const byCycle = new Map();
  for (const task of tasks) {
    const cycle = inferCycleClient(task) || 'Program';
    if (!byCycle.has(cycle)) byCycle.set(cycle, []);
    byCycle.get(cycle).push(task);
  }

  const cycleNames = [...byCycle.keys()].sort((a, b) => {
    const ai = cycleIndex.has(a) ? cycleIndex.get(a) : 999;
    const bi = cycleIndex.has(b) ? cycleIndex.get(b) : 999;
    return ai - bi || String(a).localeCompare(String(b));
  });

  const rows = [];
  for (const cycle of cycleNames) {
    const cycleTasks = byCycle.get(cycle) || [];
    rows.push(makeSyntheticRow('MODE-' + mode + '-LANE-' + cycle, cycle + ' · ' + modeLabel(mode), 'bar-lane', cycleTasks, { lane: cycle, mode }));
    for (const task of cycleTasks) {
      let cls = mspCls(task, false, !!task.cpm?.critical);
      if (mode === 'CRITICAL_TODAY') cls = task.cpm?.critical ? 'bar-critical' : 'bar-immediate';
      if (mode === 'NEAR_CRITICAL') cls = 'bar-soon';
      if (mode === 'WATCH') cls = 'bar-watch';
      if (mode === 'SOON') cls = 'bar-soon';
      if (mode === 'PAST_DUE') cls = 'bar-msp-late';
      rows.push(toGanttTask(task, cls, { cycle, relation: mode.toLowerCase() }));
    }
  }
  return rows;
}

function passesFilters(task, filterWs, filterTierVal) {
  if (!task.start || !task.finish) return false;
  if (filterWs && task.workstream !== filterWs && !task.summary) return false;
  if (filterTierVal && !task.summary) {
    const tier = tierMap.get(task.uid);
    if (tier !== filterTierVal && !task.cpm?.critical) return false;
  }
  return true;
}

function getAnchorRows(anchor, depth, filterWs, filterTierVal) {
  const prefix = String(anchor.outlineNumber || '') + '.';
  return DATA.tasks.filter((task) => {
    const outline = String(task.outlineNumber || '');
    if (task.uid !== anchor.uid && !outline.startsWith(prefix)) return false;
    if (depth !== 99 && (task.outlineLevel || 0) > (anchor.outlineLevel || 0) + depth) return false;
    return passesFilters(task, filterWs, filterTierVal);
  }).sort(compareOutline);
}

function buildCriticalFocusSet(tasks) {
  const map = new Map(tasks.map((t) => [t.uid, t]));
  const focus = new Set();
  for (const task of tasks) {
    if (!task.cpm?.critical) continue;
    focus.add(task.uid);
    (task.predecessors || []).forEach((p) => { if (map.has(p)) focus.add(p); });
    (task.successors || []).forEach((s) => { if (map.has(s)) focus.add(s); });
  }
  return focus;
}

function buildNetworkContext(baseIds, filterWs, filterTierVal, focusIds = null) {
  return DATA.tasks.filter((task) => {
    if (baseIds.has(task.uid)) return false;
    if (!passesFilters(task, filterWs, filterTierVal)) return false;
    const upstream = (task.successors || []).some((s) => baseIds.has(s));
    const downstream = (task.predecessors || []).some((p) => baseIds.has(p));
    if (!upstream && !downstream) return false;
    if (focusIds && focusIds.size) {
      const related = (task.successors || []).some((s) => focusIds.has(s)) || (task.predecessors || []).some((p) => focusIds.has(p));
      if (!related) return false;
    }
    return true;
  }).sort(compareOutline).map((task) => ({
    task,
    relation: (task.successors || []).some((s) => baseIds.has(s)) ? 'upstream' : 'downstream'
  }));
}

function findAnchorForTask(task) {
  if (!task) return null;
  const outline = String(task.outlineNumber || '');
  return laneAnchors
    .filter((anchor) => task.uid === anchor.uid || outline === anchor.outlineNumber || outline.startsWith(String(anchor.outlineNumber || '') + '.'))
    .sort((a, b) => String(b.outlineNumber || '').length - String(a.outlineNumber || '').length)[0] || null;
}

function collectRecursiveNetwork(seedUid, direction, filterWs, filterTierVal, maxNodes) {
  const out = new Map();
  const queue = [seedUid];
  const seen = new Set();
  while (queue.length && out.size < maxNodes) {
    const uid = queue.shift();
    if (seen.has(uid)) continue;
    seen.add(uid);
    const task = taskByUid.get(uid);
    if (!task) continue;
    if (uid !== seedUid && !passesFilters(task, filterWs, filterTierVal)) continue;
    if (!out.has(uid)) out.set(uid, task);
    const nextIds = direction === 'upstream' ? (task.predecessors || []) : (task.successors || []);
    nextIds.forEach((nextUid) => {
      if (!seen.has(nextUid)) queue.push(nextUid);
    });
  }
  return [...out.values()];
}

function collectDriverChain(seedTask, filterWs, filterTierVal, maxNodes) {
  const out = new Map();
  const queue = [seedTask];
  while (queue.length && out.size < maxNodes) {
    const task = queue.shift();
    if (!task || out.has(task.uid)) continue;
    if (task.uid !== seedTask.uid && !passesFilters(task, filterWs, filterTierVal)) continue;
    out.set(task.uid, task);
    const nextIds = (task.drivers && task.drivers.length ? task.drivers.map((d) => d.uid) : (task.predecessors || []));
    nextIds.forEach((uid) => {
      const nextTask = taskByUid.get(uid);
      if (nextTask && !out.has(nextTask.uid)) queue.push(nextTask);
    });
  }
  return [...out.values()].sort(compareOutline);
}

function buildFocusedMilestoneRows(focusTask, filterWs, filterTierVal) {
  const anchor = findAnchorForTask(focusTask);
  const cycle = anchor?.cycle || null;
  const track = anchor?.track || null;
  const driverChain = focusDriversOnly
    ? collectDriverChain(focusTask, filterWs, filterTierVal, 80).filter((t) => t.uid !== focusTask.uid)
    : [];
  const upstream = focusDriversOnly
    ? driverChain
    : collectRecursiveNetwork(focusTask.uid, 'upstream', filterWs, filterTierVal, 80)
      .filter((t) => t.uid !== focusTask.uid)
      .sort(compareOutline);
  const downstream = focusDriversOnly
    ? []
    : collectRecursiveNetwork(focusTask.uid, 'downstream', filterWs, filterTierVal, 40)
      .filter((t) => t.uid !== focusTask.uid)
      .sort(compareOutline);

  const rows = [];
  const bucketRows = [...upstream, focusTask, ...downstream].filter((t) => t.start && t.finish);
  if (cycle) rows.push(makeSyntheticRow('FOCUS-LANE-' + cycle, cycle + (focusDriversOnly ? ' driver chain' : ' focused network'), 'bar-lane', bucketRows, { lane: cycle }));
  if (track) rows.push(makeSyntheticRow('FOCUS-TRACK-' + focusTask.uid, (anchor?.label || track) + (focusDriversOnly ? ' — driver chain' : ' — milestone network'), 'bar-track', bucketRows, { lane: cycle, track }));
  upstream.forEach((task) => rows.push(toGanttTask(task, focusDriversOnly ? 'context-driver' : 'context-upstream', { cycle, track, relation: focusDriversOnly ? 'driver' : 'upstream' })));
  rows.push(toGanttTask(focusTask, 'bar-focus', { cycle, track, relation: 'focus' }));
  downstream.forEach((task) => rows.push(toGanttTask(task, 'context-downstream', { cycle, track, relation: 'downstream' })));
  return rows;
}

function buildGanttTasks() {
  const filterWsEl = document.getElementById('filterWs');
  const depthEl = document.getElementById('filterLv');
  const filterTierEl = document.getElementById('filterTier');
  const filterWs      = filterWsEl ? filterWsEl.value : '';
  const depth         = depthEl ? (parseInt(depthEl.value) || 2) : 2;
  const filterTierVal = filterTierEl ? filterTierEl.value : '';
  const focusTask     = focusedTaskUid ? taskByUid.get(Number(focusedTaskUid) || focusedTaskUid) : null;

  if (timelineMode === 'ALL' || timelineMode === 'CRITICAL' || timelineMode === 'DRIVER') {
    let rows = [];
    if (timelineMode === 'ALL') rows = buildAllMilestoneRows();
    else if (timelineMode === 'CRITICAL') rows = buildCriticalPathRows();
    else rows = buildDriverFlowRows();

    const visibleIds = new Set(rows
      .filter((row) => row._meta && !row._meta.synthetic && row._meta.uid != null)
      .map((row) => String(row._meta.uid)));
    for (const row of rows) {
      if (!row._meta || row._meta.synthetic || row._meta.uid == null) continue;
      row.dependencies = (row._meta.predecessorIds || [])
        .filter((p) => visibleIds.has(String(p)))
        .map(String)
        .join(',');
    }
    return rows;
  }

  if (['SLIPPING', 'PAST_DUE', 'CRITICAL_TODAY', 'NEAR_CRITICAL', 'SOON', 'WATCH'].includes(timelineMode)) {
    const modeRows = buildModeFilteredRows(timelineMode, filterWs, filterTierVal);
    const visibleIds = new Set(modeRows
      .filter((row) => row._meta && !row._meta.synthetic && row._meta.uid != null)
      .map((row) => String(row._meta.uid)));
    for (const row of modeRows) {
      if (!row._meta || row._meta.synthetic || row._meta.uid == null) continue;
      row.dependencies = (row._meta.predecessorIds || [])
        .filter((p) => visibleIds.has(String(p)))
        .map(String)
        .join(',');
    }
    return modeRows;
  }

  if (focusTask) {
    const focusedRows = buildFocusedMilestoneRows(focusTask, filterWs, filterTierVal);
    const visibleIds = new Set(focusedRows.filter((row) => row._meta && !row._meta.synthetic && row._meta.uid != null).map((row) => String(row._meta.uid)));
    for (const row of focusedRows) {
      if (!row._meta || row._meta.synthetic || row._meta.uid == null) continue;
      row.dependencies = (row._meta.predecessorIds || []).filter((p) => visibleIds.has(String(p))).map(String).join(',');
    }
    return focusedRows;
  }

  const visibleAnchors = laneAnchors
    .filter((anchor) => selectedCycles.has(anchor.cycle) && selectedTracks.has(anchor.track))
    .sort((a, b) => cycleOrder.indexOf(a.cycle) - cycleOrder.indexOf(b.cycle) || trackOrder.indexOf(a.track) - trackOrder.indexOf(b.track));

  let branchBuckets = [];
  for (const cycle of cycleOrder) {
    if (!selectedCycles.has(cycle)) continue;
    const branches = visibleAnchors
      .filter((anchor) => anchor.cycle === cycle)
      .map((anchor) => ({ anchor, rows: getAnchorRows(anchor, depth, filterWs, filterTierVal) }))
      .filter((branch) => branch.rows.length);
    if (branches.length) branchBuckets.push({ cycle, branches });
  }

  if (criticalNetworkOnly) {
    const allBranchRows = branchBuckets.flatMap((bucket) => bucket.branches.flatMap((branch) => branch.rows));
    const focusIds = buildCriticalFocusSet(allBranchRows);
    if (focusIds.size) {
      branchBuckets = branchBuckets.map((bucket) => ({
        ...bucket,
        branches: bucket.branches.map((branch) => ({
          ...branch,
          rows: branch.rows.filter((row) => row.summary || focusIds.has(row.uid))
        })).filter((branch) => branch.rows.length > 1)
      })).filter((bucket) => bucket.branches.length);
    }
  }

  const orderedRows = [];
  const baseIds = new Set();
  for (const bucket of branchBuckets) {
    const allRows = bucket.branches.flatMap((branch) => branch.rows);
    orderedRows.push(makeSyntheticRow('LANE-' + bucket.cycle, bucket.cycle + ' swimlane', 'bar-lane', allRows, { lane: bucket.cycle }));
    for (const branch of bucket.branches) {
      orderedRows.push(makeSyntheticRow('TRACK-' + branch.anchor.id, branch.anchor.label, 'bar-track', branch.rows, { lane: bucket.cycle, track: branch.anchor.track }));
      for (const row of branch.rows) {
        if (row.uid === branch.anchor.uid) continue;
        orderedRows.push(toGanttTask(row, tierCls(row.uid, row.summary, row.milestone, !!(row.cpm && row.cpm.critical)), { cycle: bucket.cycle, track: branch.anchor.track }));
        baseIds.add(row.uid);
      }
    }
  }

  if (includeContext && baseIds.size) {
    const focusIds = criticalNetworkOnly
      ? buildCriticalFocusSet([...baseIds].map((uid) => DATA.tasks.find((t) => t.uid === uid)).filter(Boolean))
      : null;
    const contextRows = buildNetworkContext(baseIds, filterWs, filterTierVal, focusIds);
    if (contextRows.length) {
      orderedRows.push(makeSyntheticRow('NETWORK-CONTEXT', 'Upstream / downstream drivers', 'bar-track', contextRows.map((x) => x.task), { track: 'Network Context' }));
      for (const entry of contextRows) {
        orderedRows.push(toGanttTask(entry.task, entry.relation === 'upstream' ? 'context-upstream' : 'context-downstream', { relation: entry.relation, track: 'Network Context' }));
      }
    }
  }

  const visibleIds = new Set(orderedRows.filter((row) => row._meta && !row._meta.synthetic && row._meta.uid != null).map((row) => String(row._meta.uid)));
  for (const row of orderedRows) {
    if (!row._meta || row._meta.synthetic || row._meta.uid == null) continue;
    row.dependencies = (row._meta.predecessorIds || []).filter((p) => visibleIds.has(String(p))).map(String).join(',');
  }

  return orderedRows;
}

function buildPopupHtml(task) {
  const m = task._meta || {};
  if (m.synthetic) {
    return '<div class="gantt-popup">'
      + '<div class="title">' + esc(task.name) + '</div>'
      + '<div class="popup-row"><span class="popup-lbl">Window</span><span>' + esc(task.start) + ' → ' + esc(task.end) + '</span></div>'
      + (m.lane ? '<div class="popup-row"><span class="popup-lbl">Lane</span><span>' + esc(m.lane) + '</span></div>' : '')
      + (m.track ? '<div class="popup-row"><span class="popup-lbl">Track</span><span>' + esc(m.track) + '</span></div>' : '')
      + '</div>';
  }
  const tierLabel = m.critical ? 'CRITICAL' : (m.tier || '—');
  const tierCls2  = m.critical ? 'pop-imm' : (m.tier === 'IMMEDIATE' ? 'pop-imm' : m.tier === 'SOON' ? 'pop-soon' : 'pop-watch');
  return \`<div class="gantt-popup">
    <div class="title">\${esc(task.name)}</div>
    <div class="popup-row"><span class="popup-lbl">WBS</span><span>\${esc(m.wbs||'—')}</span></div>
    <div class="popup-row"><span class="popup-lbl">Workstream</span><span>\${esc(m.workstream||'—')}</span></div>
    <div class="popup-row"><span class="popup-lbl">Owner</span><span>\${esc(m.resource||'Unassigned')}</span></div>
    <div class="popup-row"><span class="popup-lbl">Dates</span><span>\${esc(task.start)} → \${esc(task.end)}\${m.baselineFinish?' <span style="color:#6b7280;font-size:10px">(BL: '+m.baselineFinish+')</span>':''}</span></div>
    \${m.slipDays > 0 ? '<div class="popup-row"><span class="popup-lbl">Slip</span><span style="color:#c0392b;font-weight:700">+'+m.slipDays+' days vs baseline</span></div>' : ''}
    <div class="popup-row"><span class="popup-lbl">Priority</span><span class="\${tierCls2}">\${tierLabel}</span></div>
    \${m.whatIsWrong ? '<div style="margin-top:6px;padding-top:6px;border-top:1px solid #dde3ed"><div class="popup-row"><span class="popup-lbl">What</span><span>\${esc(m.whatIsWrong)}</span></div><div class="popup-row"><span class="popup-lbl">Why</span><span>\${esc(m.why)}</span></div><div class="popup-row"><span class="popup-lbl">Impact</span><span>\${esc(m.impact)}</span></div><div class="popup-row"><span class="popup-lbl">Action</span><span>\${esc(m.resolution)}</span></div></div>' : ''}
  </div>\`;
}

function rebuildGantt() {
  const tasks = buildGanttTasks();
  const realTaskCount = tasks.filter((t) => t._meta && !t._meta.synthetic).length;
  const focusTask = focusedTaskUid ? taskByUid.get(Number(focusedTaskUid) || focusedTaskUid) : null;
  const viewLabel = modeLabel(timelineMode);
  document.getElementById('ganttInfo').textContent = realTaskCount + ' tasks shown · Showing: ' + viewLabel + ' · ' + (focusTask ? ((focusDriversOnly ? 'Driver chain: ' : 'Focused milestone: ') + focusTask.name) : ([...selectedCycles].join(', ') + ' · ' + [...selectedTracks].join(', ')));
  if (!tasks.length) {
    document.getElementById('ganttWrap').innerHTML = '<div style="padding:32px;text-align:center;color:var(--muted)">No tasks match current filters.</div>';
    ganttChart = null;
    return;
  }
  // Reset svg target
  document.getElementById('ganttWrap').innerHTML = '<svg id="gantt"></svg>';
  const viewMode = document.getElementById('viewMode').value;
  ganttChart = new Gantt('#gantt', tasks, {
    view_modes: buildViewModes(viewMode),
    bar_height: 20,
    bar_corner_radius: 3,
    padding: 14,
    scroll_to: smartScrollDate || 'today',
    readonly: true,
    today_button: false,
    show_expected_progress: false,
    lines: 'both',
    popup: (task) => buildPopupHtml(task),
    on_click: (task) => {
      let uid = task._meta && task._meta.uid;
      if (!uid && (task.id === 'CP-ROW' || task.id === 'DRV-ROW')) {
        const firstCritical = DATA.tasks
          .filter((t) => t.cpm?.critical && !t.summary && t.active)
          .sort((a, b) => (a.cpm?.earlyStart ?? 0) - (b.cpm?.earlyStart ?? 0))[0];
        uid = firstCritical?.uid;
      }
      if (!uid) return;
      openNetworkCanvas(uid);
      const tier = tierMap.get(uid);
      if (!tier) return;
      switchTier(tier);
      setTimeout(() => {
        const card = document.getElementById('card-' + uid);
        if (card) { card.classList.add('expanded'); card.scrollIntoView({ behavior: 'smooth', block: 'nearest' }); }
      }, 60);
    }
  });
  paintSwimlaneBands(tasks);
}

function scrollFocus() {
  if (!ganttChart) return;
  if (smartScrollDate) {
    if (typeof ganttChart.set_scroll_position === 'function') {
      ganttChart.set_scroll_position(smartScrollDate);
      return;
    }
    if (typeof ganttChart.scroll_to === 'function') {
      ganttChart.scroll_to(smartScrollDate);
      return;
    }
  }
  if (typeof ganttChart.scroll_current === 'function') {
    ganttChart.scroll_current();
  }
}

function changeView() {
  if (!ganttChart) return;
  ganttChart.change_view_mode(document.getElementById('viewMode').value);
  refreshControlTooltips();
  paintSwimlaneBands();
}

function scrollToday() {
  if (!ganttChart) return;
  if (typeof ganttChart.set_scroll_position === 'function') {
    ganttChart.set_scroll_position('today');
    return;
  }
  if (typeof ganttChart.scroll_current === 'function') {
    ganttChart.scroll_current();
  }
}

// Initial render
rebuildGantt();

// ── Workstream Health Strip ─────────────────────────────────────────────────
function renderWsStrip() {
  const health = m.workstreamHealth || [];
  const wsEl = document.getElementById('wsStrip');
  wsEl.innerHTML = health.slice(0, 30).map(ws => {
    const total = ws.total || 1;
    const iPct = Math.round(ws.immediate / total * 100);
    const sPct = Math.round(ws.soon / total * 100);
    const wPct = 100 - iPct - sPct;
    return \`<div class="ws-chip" data-ws="\${esc(ws.workstream)}" onclick="filterByWs('\${esc(ws.workstream)}', this)" title="\${esc(ws.workstream)}">
      <div class="ws-name">\${esc(ws.workstream)}</div>
      <div class="ws-bars">
        <span style="width:\${iPct}%;background:var(--red)"></span>
        <span style="width:\${sPct}%;background:var(--amber)"></span>
        <span style="width:\${wPct}%;background:var(--blue)"></span>
      </div>
      <div class="ws-counts">\${ws.total} tasks · \${ws.immediate>0?'<span style="color:var(--red)">'+ws.immediate+' imm</span> · ':''}\${ws.soon>0?ws.soon+' soon · ':''}\${ws.watch} watch</div>
    </div>\`;
  }).join('');
}

function filterByWs(ws, sourceEl) {
  document.getElementById('filterWs').value = ws;
  document.querySelectorAll('.ws-chip').forEach(c => c.classList.remove('active'));
  if (sourceEl) {
    sourceEl.classList.add('active');
  } else {
    const match = document.querySelector('.ws-chip[data-ws="' + ws.replace(/"/g, '\\"') + '"]');
    if (match) match.classList.add('active');
  }
  rebuildGantt();
}

renderWsStrip();
refreshViewFilterTooltip();
refreshControlTooltips();
</script>
</body>
</html>`;
}

function main() {
  ensureDir(DOCS_DIR);
  const taskCsv = pickPrimaryTaskCsv();
  const supplemental = parseSupplementalPlan(findOptionalStagingFile(SUPPLEMENTAL_PLAN_NAME));
  const resourceTable = parseResourceTable(findOptionalStagingFile(RESOURCE_TABLE_NAME));
  const projectTeam = parseProjectTeamRoster(findOptionalStagingFile(PROJECT_TEAM_NAME));
  const tasks = loadTasks(taskCsv);
  mergeSupplementalIntoBaseRows(tasks, supplemental);
  const enrichment = enrichWorkingTasks(tasks, projectTeam, resourceTable);
  normalizeTaskWorkstreams(tasks);
  const cpmSummary = buildCpm(tasks);
  const metrics = computeScheduleStats(tasks, cpmSummary);
  const mermaid = buildMermaid(tasks);
  const laneAnchors = buildLaneAnchors(tasks);

  const actionItems = [
    ...(metrics.immediate || []).map(t => ({
      taskId: t.uid,
      attentionLane: 'Act Now',
      severity: (t.cpm && t.cpm.critical) ? 'Critical' : (t.slipDays >= 30 ? 'High' : 'Medium'),
      type: (t.cpm && t.cpm.critical) ? 'Critical Path' : 'Schedule Risk',
      category: t.workstream || '',
      name: t.whatIsWrong || t.name,
      title: t.name,
      milestoneDate: t.finish
    })),
    ...(metrics.soon || []).map(t => ({
      taskId: t.uid,
      attentionLane: 'At Risk Soon',
      severity: t.slipDays >= 30 ? 'High' : 'Medium',
      type: 'Schedule Risk',
      category: t.workstream || '',
      name: t.whatIsWrong || t.name,
      title: t.name,
      milestoneDate: t.finish
    })),
    ...(metrics.watch || []).map(t => ({
      taskId: t.uid,
      attentionLane: 'Watchlist',
      severity: (t.cpm && t.cpm.totalSlack != null && t.cpm.totalSlack <= 2) ? 'High' : 'Low',
      type: 'Monitor',
      category: t.workstream || '',
      name: t.whatIsWrong || t.name,
      title: t.name,
      milestoneDate: t.finish
    }))
  ];

  const payload = {
    generatedAt: new Date().toISOString(),
    source: { file: path.basename(taskCsv) },
    enrichment,
    metrics,
    mermaid,
    laneAnchors,
    actionItems,
    tasks: tasks.sort((a, b) => (a.outlineNumber || '').localeCompare(b.outlineNumber || '') || a.uid - b.uid)
  };

  fs.writeFileSync(OUT_JSON, JSON.stringify(payload, null, 2), "utf8");
  fs.writeFileSync(OUT_MERMAID, mermaid, "utf8");
  fs.writeFileSync(OUT_HTML, renderHtml(payload), "utf8");

  console.log("Schedule intelligence outputs generated:");
  console.log(`  HTML: ${OUT_HTML}`);
  console.log(`  Data: ${OUT_JSON}`);
  console.log(`  Mermaid: ${OUT_MERMAID}`);
  console.log(`  Enrichment -> Resource: ${enrichment.resourceRowsUpdated}, BPO: ${enrichment.bpoRowsUpdated}, Workstream: ${enrichment.workstreamRowsUpdated}`);
  console.log(`  CPM -> Duration: ${metrics.cpmSummary.projectDuration}d, Critical tasks: ${metrics.criticalTasks}, Cyclic: ${metrics.cpmSummary.cyclic}`);
}

main();

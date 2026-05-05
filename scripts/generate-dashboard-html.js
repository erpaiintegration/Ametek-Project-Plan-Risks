/* eslint-disable no-console */
require("dotenv/config");
const { Client } = require("@notionhq/client");
const fs = require("fs");
const path = require("path");

// Inline Chart.js so the HTML works offline / file:// / Notion embed
const CHARTJS_PATH = path.join(__dirname, "..", "node_modules", "chart.js", "dist", "chart.umd.js");
const CHARTJS_INLINE = fs.existsSync(CHARTJS_PATH) ? fs.readFileSync(CHARTJS_PATH, "utf8") : null;
const PLOTLY_PATH = path.join(__dirname, "..", "node_modules", "plotly.js-dist-min", "plotly.min.js");
const PLOTLY_INLINE = fs.existsSync(PLOTLY_PATH) ? fs.readFileSync(PLOTLY_PATH, "utf8") : null;
const INLINE_LIBS = process.env.DASHBOARD_INLINE_LIBS === "1";
const notion = new Client({ auth: process.env.NOTION_API_KEY });
const TASKS_DB_ID = process.env.NOTION_TASKS_DB_ID;
const RISKS_DB_ID = process.env.NOTION_RISKS_DB_ID || "357ae9be-8a60-8190-b720-c130c7104cf1";
const OUT_FILE = path.join(__dirname, "..", "dashboard.html");

function text(parts) { return (parts || []).map(p => p.plain_text).join(""); }
function titleProp(p) { return p?.type === "title" ? text(p.title || []) : ""; }
function rich(p) { return p?.type === "rich_text" ? text(p.rich_text || []) : ""; }
function dateVal(p) { return p?.type === "date" ? p.date?.start || null : null; }
function numVal(p) { return p?.type === "number" ? p.number : null; }
function checkVal(p) { return p?.type === "checkbox" ? !!p.checkbox : false; }
function selVal(p) {
  if (p?.type === "select") return p.select?.name || "";
  if (p?.type === "status") return p.status?.name || "";
  return "";
}
function relIds(p) { return p?.type === "relation" ? (p.relation || []).map(r => r.id) : []; }
function relCount(p) { return relIds(p).length; }
function textLike(p) {
  if (!p) return "";
  if (p.type === "rich_text") return text(p.rich_text || []);
  if (p.type === "title") return text(p.title || []);
  if (p.type === "number") return p.number == null ? "" : String(p.number);
  if (p.type === "select") return p.select?.name || "";
  if (p.type === "status") return p.status?.name || "";
  if (p.type === "formula") {
    const f = p.formula || {};
    if (typeof f.string === "string") return f.string;
    if (typeof f.number === "number") return String(f.number);
  }
  return "";
}
function parseIntOrNull(v) {
  const n = parseInt(String(v || "").trim(), 10);
  return Number.isFinite(n) ? n : null;
}
function extractOutlineLevel(props) {
  const direct = [
    props["Outline Level"],
    props["Level"],
    props["WBS Level"],
    props["Task Level"]
  ];
  for (const p of direct) {
    const n = p?.type === "number" ? p.number : parseIntOrNull(textLike(p));
    if (n != null) return n;
  }
  const outlineNumber = textLike(props["Outline Number"]) || textLike(props["WBS"]);
  if (outlineNumber) {
    const levelFromWbs = outlineNumber.split(".").filter(Boolean).length;
    if (levelFromWbs > 0) return levelFromWbs;
  }
  return null;
}
function extractParentUid(props) {
  const candidates = [
    props["Parent Unique ID"],
    props["Parent UID"],
    props["Parent Task ID"],
    props["Parent Task Unique ID"],
    props["Summary Task ID"]
  ];
  for (const p of candidates) {
    const n = p?.type === "number" ? p.number : parseIntOrNull(textLike(p));
    if (n != null) return n;
  }
  return null;
}
function splitNames(v) {
  return [...new Set(String(v || "").split(/[;|,/]/).map(x => x.trim()).filter(Boolean))];
}
function relCountByNames(props, names) {
  for (const n of names) { const p = props[n]; if (p?.type === "relation") return relCount(p); }
  return 0;
}

function simpleSentence(s) {
  return String(s || "")
    .replace(/\s+/g, " ")
    .replace(/[_-]+/g, " ")
    .trim()
    .replace(/^\w/, c => c.toUpperCase())
    .replace(/[.;:!?]+$/, "");
}

function copilotCategory(rec) {
  const seed = `${rec.name} ${rec.category} ${rec.type}`.toLowerCase();
  if (/(uat|test|defect|qa|validation)/.test(seed)) return "Testing & Quality";
  if (/(schedule|delay|milestone|late|slip)/.test(seed)) return "Schedule & Milestones";
  if (/(data|migration|load|master|interface)/.test(seed)) return "Data & Integration";
  if (/(security|access|grc|auth|firefighter)/.test(seed)) return "Security & Access";
  if (/(cutover|go-live|deployment|production)/.test(seed)) return "Cutover & Go-Live";
  if (/(resource|owner|capacity|staff)/.test(seed)) return "Resourcing & Ownership";
  return rec.category && rec.category !== "(Uncategorized)" ? rec.category : "Execution Risk";
}

function copilotImpact(severity, linkedCount, type) {
  const sev = String(severity || "").toLowerCase();
  const base = /issue/i.test(type)
    ? "This is blocking work right now"
    : "This can turn into a blocker if not handled";
  if (/critical|very high|urgent|severe|high/.test(sev)) {
    return `${base}, and it could delay key milestones soon.`;
  }
  if (linkedCount >= 4) {
    return `${base}, and it affects multiple dependent tasks.`;
  }
  return `${base}, with moderate schedule impact if left open.`;
}

function copilotActions(rec, linkedTasks) {
  const names = linkedTasks.slice(0, 3).map(t => t.name);
  const top = names.length ? `Target first: ${names.join(", ")}.` : "Target the nearest open milestone first.";
  return [
    "Assign one owner and due date today.",
    "Run a 15-minute root-cause check with impacted workstream leads.",
    top
  ];
}

function fmtIso(d) {
  if (!d) return "(date missing)";
  return String(d).slice(0, 10);
}

function rankLinkedTask(t) {
  if (!t) return -1;
  let score = 0;
  if (t.isSlipped) score += 40;
  if (t.isOverdueStart) score += 25;
  if (t.isDue14) score += 15;
  if (t.pct < 100) score += 10;
  if ((t.sucCount || 0) > 0) score += 8;
  if ((t.predCount || 0) > 0) score += 5;
  if (t.milestone) score += 6;
  score += Math.min(10, t.slipDays || 0);
  return score;
}

function pickPrimaryTask(linkedTasks) {
  if (!linkedTasks.length) return null;
  return [...linkedTasks].sort((a, b) => rankLinkedTask(b) - rankLinkedTask(a))[0] || null;
}

function pickImpactedMilestone(task, taskById) {
  if (!task || !taskById) return null;
  const successors = (task.sucIds || []).map(id => taskById.get(id)).filter(Boolean);
  const directMilestone = successors.find(s => s.milestone) || null;
  if (directMilestone) return directMilestone;
  return successors
    .filter(s => s.finish || s.start)
    .sort((a, b) => String(a.finish || a.start).localeCompare(String(b.finish || b.start)))[0] || null;
}

function specificIssueNarrative(rec, linkedTasks, taskById) {
  const t = pickPrimaryTask(linkedTasks);
  if (!t) return simpleSentence(rec.name);

  const startDate = fmtIso(t.start);
  const finishDate = fmtIso(t.finish);
  const milestone = pickImpactedMilestone(t, taskById);
  const ws = t.workstream || "This workstream";
  const status = t.pct >= 100 ? "is complete" : `is only ${t.pct || 0}% complete`;

  let sentence = `${ws} is scheduled to start on ${startDate}, but task \"${simpleSentence(t.name)}\" ${status} and is planned to finish ${finishDate}.`;
  if (milestone) {
    sentence += ` This creates a scheduling conflict and impacts milestone \"${simpleSentence(milestone.name)}\" (${fmtIso(milestone.finish || milestone.start)}).`;
  } else {
    sentence += " This creates a scheduling conflict for downstream activities.";
  }
  return sentence;
}

function specificImpactNarrative(rec, linkedTasks, taskById) {
  const t = pickPrimaryTask(linkedTasks);
  if (!t) return copilotImpact(rec.severity, linkedTasks.length, rec.type);

  const milestone = pickImpactedMilestone(t, taskById);
  const slip = t.slipDays != null ? ` by approximately ${t.slipDays} day(s)` : "";
  const predTxt = (t.predCount || 0) > 0 ? `${t.predCount} predecessor(s)` : "no predecessor constraints";
  const sucTxt = (t.sucCount || 0) > 0 ? `${t.sucCount} successor(s)` : "limited downstream links";

  if (milestone) {
    return `Because \"${simpleSentence(t.name)}\" is not complete, milestone \"${simpleSentence(milestone.name)}\" can be pushed out${slip}. Dependency profile: ${predTxt}, ${sucTxt}.`;
  }
  return `Because \"${simpleSentence(t.name)}\" is not complete, dependent schedule dates are at risk${slip}. Dependency profile: ${predTxt}, ${sucTxt}.`;
}

function specificActions(rec, linkedTasks, taskById) {
  const t = pickPrimaryTask(linkedTasks);
  if (!t) return copilotActions(rec, linkedTasks);

  const milestone = pickImpactedMilestone(t, taskById);
  const baseline = fmtIso(t.baselineFinish || t.finish || t.start);
  const currentFinish = fmtIso(t.finish || t.start);
  const finishTarget = t.baselineFinish ? fmtIso(t.baselineFinish) : "baseline target";
  const actions = [];

  if ((t.predCount || 0) > 0) {
    actions.push(`Review predecessor links for \"${simpleSentence(t.name)}\": remove or relax any non-driving predecessor causing late start (governance approval required).`);
  } else {
    actions.push(`Resequence \"${simpleSentence(t.name)}\" to start earlier than ${fmtIso(t.start)} by pulling preparatory steps forward.`);
  }

  actions.push(`Change duration for \"${simpleSentence(t.name)}\" by fast-tracking: compress from current plan ending ${currentFinish} toward ${finishTarget} (add temporary owner support / parallelize checks).`);

  if (milestone) {
    actions.push(`Protect milestone \"${simpleSentence(milestone.name)}\" (${fmtIso(milestone.finish || milestone.start)}): run a daily dependency check and re-baseline only if recovery to ${baseline} is not feasible.`);
  } else {
    actions.push(`Run a downstream impact check on ${t.sucCount || 0} successor task(s) and update handoff dates before end of day.`);
  }

  return actions;
}

function buildLinkageImpact(primaryTask, impactedMilestone) {
  if (!primaryTask) return "No linked task context available.";
  const pred = primaryTask.predCount || 0;
  const suc = primaryTask.sucCount || 0;
  const base = `${pred} predecessor(s) → ${simpleSentence(primaryTask.name)} → ${suc} successor(s)`;
  if (impactedMilestone) {
    return `${base}; direct milestone impact: ${simpleSentence(impactedMilestone.name)} on ${fmtIso(impactedMilestone.finish || impactedMilestone.start)}.`;
  }
  return `${base}; downstream successors are exposed to schedule slippage.`;
}

function normalizeToken(s) {
  return String(s || "").toLowerCase().replace(/[^a-z0-9]+/g, " ").trim();
}

function inferLinkedTasksFromText(rec, tasks, workstreamHints) {
  const seed = normalizeToken(`${rec.name} ${rec.category} ${rec.type}`);
  let pool = tasks.filter(t => t.pct < 100);

  const matchedHints = workstreamHints.filter(ws => {
    const n = normalizeToken(ws);
    return n && seed.includes(n);
  });
  if (matchedHints.length) {
    pool = pool.filter(t => matchedHints.some(ws => String(t.workstream || "").toLowerCase().includes(String(ws).toLowerCase())));
  }

  if (!pool.length) return [];

  const keyTaskRegex = /(post load validation|pre load validation|simulate load|load file|error and defect mitigation|load)/i;
  const scored = pool.map(t => {
    let score = 0;
    if (t.isSlipped) score += 30;
    if (t.isOverdueStart) score += 20;
    if (t.isDue14) score += 10;
    if (keyTaskRegex.test(t.name || "")) score += 12;
    if ((t.sucCount || 0) > 0) score += 6;
    if ((t.predCount || 0) > 0) score += 4;
    score += Math.min(10, t.slipDays || 0);
    return { t, score };
  }).sort((a, b) => b.score - a.score || (b.t.slipDays || 0) - (a.t.slipDays || 0));

  return scored.slice(0, 3).map(x => x.t);
}

function daysDiff(a, b) {
  if (!a || !b) return null;
  const da = new Date(a), db = new Date(b);
  if (isNaN(da) || isNaN(db)) return null;
  return Math.round((da - db) / 86400000);
}

async function fetchAll(dbId, label) {
  const out = []; let cursor;
  do {
    const r = await notion.databases.query({ database_id: dbId, start_cursor: cursor, page_size: 100 });
    out.push(...r.results.filter(x => x.object === "page"));
    cursor = r.has_more ? r.next_cursor : undefined;
  } while (cursor);
  console.log(`${label}: ${out.length} rows`);
  return out;
}

async function main() {
  const [rawTasks, rawRisks] = await Promise.all([
    fetchAll(TASKS_DB_ID, "tasks"),
    fetchAll(RISKS_DB_ID, "risks")
  ]);

  const now = new Date();
  const in14 = new Date(now.getTime() + 14 * 86400000);
  const in30 = new Date(now.getTime() + 30 * 86400000);

  // ── Process risks/issues ───────────────────────────────────────────────────
  const riskRecords = [];
  const taskIdToRisks = new Map(); // taskPageId → [{ name, type, severity, category }]

  for (const r of rawRisks) {
    const p = r.properties || {};
    const status = selVal(p.Status) || "";
    if (/resolved|closed/i.test(status)) continue;

    const type = selVal(p.Type) || "Risk";
    const category = selVal(p["Risk Type"]) || selVal(p.Category) || selVal(p.Family) || "(Uncategorized)";
    const severity = selVal(p.Severity) || selVal(p.Priority) || selVal(p.Impact) || "(Unrated)";
    const name = titleProp(p["Name"]) || titleProp(p["Title"]) ||
      Object.values(p).find(x => x?.type === "title") ?
        text((Object.values(p).find(x => x?.type === "title") || {}).title || []) : "(Untitled)";

    const linkedTaskIds = [
      ...relIds(p.Tasks), ...relIds(p["Related Tasks"]),
      ...relIds(p["Impacted Tasks"]), ...relIds(p.Task), ...relIds(p["Task Links"])
    ];

    const rec = { id: r.id, name, type, category, severity, status, linkedTaskIds, typeKey: `${type} — ${category}` };
    riskRecords.push(rec);

    for (const tid of linkedTaskIds) {
      if (!taskIdToRisks.has(tid)) taskIdToRisks.set(tid, []);
      taskIdToRisks.get(tid).push({ name, type, severity, category });
    }
  }

  // Risk type breakdown
  const riskTypeMap = new Map();
  for (const rec of riskRecords) {
    riskTypeMap.set(rec.typeKey, (riskTypeMap.get(rec.typeKey) || 0) + 1);
  }
  const riskTypeBreakdown = [...riskTypeMap.entries()]
    .sort((a, b) => b[1] - a[1])
    .map(([typeKey, count]) => ({ typeKey, count }));

  // ── Process tasks ──────────────────────────────────────────────────────────
  const tasks = [];
  let totalOpen = 0, totalDone = 0, slippedOpen = 0, overdueStarts = 0, due14Count = 0;
  const workstreamPressure = new Map();
  const statusCounts = new Map();

  for (const t of rawTasks) {
    const p = t.properties || {};
    const uid = numVal(p["Unique ID"]);
    const name = titleProp(p["Task Name"]) || "(Untitled)";
    const workstream = rich(p.Workstream).trim() || "(Unassigned)";
    const start = dateVal(p.Start);
    const finish = dateVal(p.Finish);
    const baselineFinish = dateVal(p["Baseline Finish"]) || dateVal(p["Baseline10 Finish"]);
    const pct = numVal(p["% Complete"]) ?? 0;
    const milestone = checkVal(p.Milestone);
    const outlineLevel = extractOutlineLevel(p);
    const parentUid = extractParentUid(p);
    const predCount = relCountByNames(p, ["Predecessor Tasks", "Predecessors"]);
    const sucCount = relCountByNames(p, ["Successor Tasks", "Successors"]);
    const predIds = [...relIds(p["Predecessor Tasks"]), ...relIds(p.Predecessors)];
    const sucIds = [...relIds(p["Successor Tasks"]), ...relIds(p.Successors)];
    const resourceNames = splitNames(rich(p["Resource Names"]));
    const businessOwners = splitNames(rich(p["Business Validation Owner"]));
    const assignedTo = resourceNames.length ? resourceNames[0] : (businessOwners[0] || "Unassigned");

    const statusRaw = selVal(p["Status Update"]) || selVal(p.Status) || (pct >= 100 ? "Done" : "Not started");
    statusCounts.set(statusRaw, (statusCounts.get(statusRaw) || 0) + 1);

    const isOpen = pct < 100;
    if (isOpen) totalOpen++; else totalDone++;

    const slipDays = daysDiff(finish, baselineFinish);
    const isSlipped = isOpen && slipDays != null && slipDays > 0;
    if (isSlipped) {
      slippedOpen++;
      workstreamPressure.set(workstream, (workstreamPressure.get(workstream) || 0) + 1);
    }

    const isOverdueStart = pct === 0 && start && new Date(start) < now;
    if (isOverdueStart) overdueStarts++;

    const isDue14 = isOpen && finish && new Date(finish) >= now && new Date(finish) <= in14;
    if (isDue14) due14Count++;

    const linkedRisks = taskIdToRisks.get(t.id) || [];

    tasks.push({
      id: t.id,
      uid,
      name,
      workstream,
      start,
      finish,
      baselineFinish,
      pct,
      milestone,
      outlineLevel,
      parentUid,
      predCount,
      sucCount,
      predIds,
      sucIds,
      assignedTo,
      slipDays: isSlipped ? slipDays : null,
      isSlipped,
      isOverdueStart,
      isDue14,
      status: statusRaw,
      linkedRisks
    });
  }

  const total = tasks.length;
  const slipRatePct = totalOpen ? ((slippedOpen / totalOpen) * 100).toFixed(1) : "0.0";
  const openIssues = riskRecords.filter(r => /issue/i.test(r.type));
  const openRisks = riskRecords.filter(r => !/issue/i.test(r.type));

  const topWorkstreams = [...workstreamPressure.entries()]
    .sort((a, b) => b[1] - a[1]).slice(0, 8)
    .map(([ws, cnt]) => ({ ws, cnt }));

  const statusBreakdown = [...statusCounts.entries()]
    .sort((a, b) => b[1] - a[1])
    .map(([s, c]) => ({ s, c }));

  const healthRed = slippedOpen >= 1200 || parseFloat(slipRatePct) >= 75 || due14Count >= 100 || overdueStarts >= 100;
  const healthAmber = slippedOpen >= 700 || parseFloat(slipRatePct) >= 50 || due14Count >= 60 || overdueStarts >= 40;
  const healthLabel = healthRed ? "🔴 Red" : healthAmber ? "🟠 Amber" : "🟢 Green";

  const taskById = new Map(tasks.map(t => [t.id, t]));
  const workstreamHints = [...new Set(tasks.map(t => t.workstream).filter(Boolean))]
    .flatMap(ws => String(ws).split(/[;|,/]/).map(x => x.trim()).filter(Boolean))
    .filter(ws => ws.length <= 20);
  const actionItems = riskRecords
    .map(r => {
      const explicitLinkedTasks = r.linkedTaskIds.map(id => taskById.get(id)).filter(Boolean);
      const linkedTasks = explicitLinkedTasks.length
        ? explicitLinkedTasks
        : inferLinkedTasksFromText(r, tasks, workstreamHints);
      const primaryTask = pickPrimaryTask(linkedTasks);
      const impactedMilestone = pickImpactedMilestone(primaryTask, taskById);
      const aiCategory = copilotCategory(r);
      const plainIssue = specificIssueNarrative(r, linkedTasks, taskById);
      const plainImpact = specificImpactNarrative(r, linkedTasks, taskById);
      const actions = specificActions(r, linkedTasks, taskById);
      const linkageImpact = buildLinkageImpact(primaryTask, impactedMilestone);
      const milestoneHits = linkedTasks.filter(t => t.milestone).slice(0, 3);
      const milestoneOwners = milestoneHits.map(t => ({ name: t.name, assignedTo: t.assignedTo, finish: t.finish || t.start }));
      return {
        id: r.id,
        type: r.type,
        severity: r.severity,
        status: r.status,
        sourceCategory: r.category,
        aiCategory,
        plainIssue,
        plainImpact,
        actions,
        workstream: primaryTask?.workstream || "(Unassigned)",
        owner: primaryTask?.assignedTo || "Unassigned",
        taskId: primaryTask?.id || null,
        taskName: primaryTask?.name || "(No linked task)",
        predCount: primaryTask?.predCount || 0,
        sucCount: primaryTask?.sucCount || 0,
        taskStart: primaryTask?.start || null,
        taskFinish: primaryTask?.finish || null,
        milestoneName: impactedMilestone?.name || null,
        milestoneDate: impactedMilestone?.finish || impactedMilestone?.start || null,
        linkageImpact,
        linkedTaskCount: linkedTasks.length,
        milestoneOwners
      };
    })
    .sort((a, b) => {
      const score = v => {
        const s = String(v.severity || "").toLowerCase();
        if (/(critical|very high|urgent|severe)/.test(s)) return 4;
        if (/high/.test(s)) return 3;
        if (/medium|med/.test(s)) return 2;
        if (/low/.test(s)) return 1;
        return 0;
      };
      return score(b) - score(a) || b.linkedTaskCount - a.linkedTaskCount;
    });

  const payload = {
    generatedAt: now.toISOString(),
    metrics: { total, totalOpen, totalDone, slippedOpen, slipRatePct, overdueStarts, due14: due14Count, openIssues: openIssues.length, openRisks: openRisks.length, healthLabel },
    riskTypeBreakdown,
    statusBreakdown,
    topWorkstreams,
    tasks,
    riskRecords,
    actionItems
  };

  const html = buildHtml(payload);
  fs.writeFileSync(OUT_FILE, html, "utf8");
  console.log(`\nDashboard written to: ${OUT_FILE}`);
  console.log(`Open it in a browser, then embed the hosted URL in Notion via /embed`);
}

function buildHtml(payload) {
  const data = JSON.stringify(payload).replace(/<\/script>/gi, "<\\/script>");
  const workstreams = [...new Set(payload.tasks.map(t => t.workstream))].sort();
  const wsOptions = workstreams.map(w => `<option value="${w.replace(/"/g,'&quot;')}">${w}</option>`).join('\n          ');
  const riskTypeOptions = payload.riskTypeBreakdown
    .map(r => `<option value="${r.typeKey.replace(/"/g, "&quot;")}">${r.typeKey} (${r.count})</option>`)
    .join("\n          ");
  const chartScript = INLINE_LIBS && CHARTJS_INLINE
    ? `<script>${CHARTJS_INLINE.replace(/<\/script>/gi, "<\\/script>")}<\/script>`
    : `<script src="https://cdn.jsdelivr.net/npm/chart.js@4.4.3/dist/chart.umd.min.js"><\/script>`;
  const plotlyScript = ''; // Plotly loaded lazily on first Gantt/Plan tab open

  return `<!DOCTYPE html>
<html lang="en">
<head>
<meta charset="UTF-8"/>
<meta name="viewport" content="width=device-width,initial-scale=1"/>
<title>Ametek SAP S4 — Schedule Dashboard</title>
${chartScript}
${plotlyScript}
<style>
  :root{--bg:#f4f7fb;--surface:#ffffff;--surface2:#eef3f8;--border:#d7e0ea;--text:#102033;--text2:#5f7185;--red:#d85b72;--amber:#c7922a;--green:#2f8f6b;--blue:#4b6bfb;--accent:#7b8cff;--shadow:0 10px 28px rgba(15,23,42,.06);}
  *{box-sizing:border-box;margin:0;padding:0;}
  body{background:var(--bg);color:var(--text);font-family:-apple-system,BlinkMacSystemFont,"Segoe UI",sans-serif;font-size:13px;height:100vh;display:flex;flex-direction:column;overflow:hidden;}
  header{background:var(--surface);border-bottom:1px solid var(--border);padding:10px 18px;display:flex;align-items:center;gap:12px;flex-shrink:0;box-shadow:var(--shadow);}
  header h1{font-size:15px;font-weight:700;}
  .health-badge{padding:4px 10px;border-radius:999px;font-size:12px;font-weight:700;background:#fdf0f2;color:#b9384f;border:1px solid #f4c6d0;}
  header .gen{font-size:11px;color:var(--text2);margin-left:auto;}
  .kpi-strip{display:flex;gap:8px;padding:8px 18px;flex-shrink:0;overflow-x:auto;}
  .kpi{background:var(--surface);border:1px solid var(--border);border-radius:10px;padding:9px 14px;white-space:nowrap;display:flex;align-items:baseline;gap:8px;box-shadow:var(--shadow);}
  .kpi .val{font-size:20px;font-weight:700;line-height:1;}
  .kpi .lbl{font-size:10px;color:var(--text2);text-transform:uppercase;letter-spacing:.4px;}
  .kpi.red .val{color:var(--red);}.kpi.amber .val{color:var(--amber);}.kpi.green .val{color:var(--green);}.kpi.blue .val{color:var(--blue);}
  .tab-bar{display:flex;padding:0 18px;flex-shrink:0;border-bottom:1px solid var(--border);}
  .tab-btn{padding:7px 18px;font-size:12px;font-weight:500;cursor:pointer;border:none;background:transparent;color:var(--text2);border-bottom:2px solid transparent;margin-bottom:-1px;transition:color .12s,border-color .12s;}
  .tab-btn:hover{color:var(--text);} .tab-btn.active{color:var(--blue);border-bottom-color:var(--blue);}
  .charts-strip{display:grid;grid-template-columns:220px 1fr 200px;gap:8px;padding:8px 18px;flex-shrink:0;height:165px;}
  .panel{background:var(--surface);border:1px solid var(--border);border-radius:12px;padding:10px;display:flex;flex-direction:column;overflow:hidden;box-shadow:var(--shadow);}
  .panel-title{font-size:10px;text-transform:uppercase;letter-spacing:.5px;color:var(--text2);margin-bottom:6px;flex-shrink:0;}
  .chart-wrap{flex:1;position:relative;min-height:0;}
  .mini-milestone{width:100%;height:100%;overflow:hidden;}
  .mini-milestone svg{width:100%;height:100%;display:block;}
  .type-list{display:flex;flex-direction:column;gap:4px;overflow-y:auto;flex:1;}
  .type-tile{display:flex;align-items:center;justify-content:space-between;background:var(--surface2);border:1px solid var(--border);border-radius:6px;padding:5px 10px;cursor:pointer;transition:border-color .12s,background .12s;user-select:none;}
  .type-tile:hover{border-color:var(--accent);} .type-tile.active{border-color:var(--blue);background:#eef2ff;}
  .type-tile .type-label{font-size:12px;} .type-tile .type-badge{background:var(--accent);color:#fff;border-radius:999px;padding:1px 7px;font-size:11px;font-weight:600;}
  .main-area{flex:1;display:flex;overflow:hidden;padding:0 18px 10px;gap:8px;}
  .task-section{flex:1;display:flex;flex-direction:column;overflow:hidden;}
  .section-bar{display:flex;align-items:center;gap:10px;padding:6px 0 5px;flex-shrink:0;}
  .section-bar h3{font-size:12px;font-weight:600;}
  .filter-info{font-size:11px;color:var(--text2);}
  .filter-clear{font-size:11px;color:var(--blue);cursor:pointer;text-decoration:underline;}
  .tbl-wrap{flex:1;overflow-y:auto;border:1px solid var(--border);border-radius:12px;background:var(--surface);box-shadow:var(--shadow);}
  table{width:100%;border-collapse:collapse;}
  thead th{position:sticky;top:0;background:var(--surface2);padding:6px 8px;text-align:left;font-size:10px;text-transform:uppercase;letter-spacing:.4px;color:var(--text2);border-bottom:1px solid var(--border);white-space:nowrap;}
  tbody tr{border-bottom:1px solid var(--border);cursor:pointer;}
  tbody tr:hover{background:var(--surface2);}
  tbody tr.selected{background:#eef2ff;border-left:2px solid var(--blue);}
  tbody td{padding:5px 8px;font-size:12px;vertical-align:top;}
  .pill{display:inline-block;padding:1px 7px;border-radius:999px;font-size:10px;font-weight:500;margin:1px 2px;}
  .pill.risk{background:#fff0f3;color:#bf4b63;} .pill.issue{background:#eef6ff;color:#2f67c8;}
  .pill.slip{background:#fff6e6;color:#b97712;} .pill.done{background:#ecfdf5;color:#1f7a59;}
  .t-name{font-weight:500;max-width:240px;} .t-ws{font-size:10px;color:var(--text2);margin-top:1px;}
  .empty{text-align:center;padding:30px;color:var(--text2);}
  .dep-panel{width:0;flex-shrink:0;background:var(--surface);border:1px solid var(--border);border-radius:12px;overflow:hidden;transition:width .2s ease;display:flex;flex-direction:column;box-shadow:var(--shadow);}
  .dep-panel.open{width:420px;}
  .dep-header{padding:10px 14px;border-bottom:1px solid var(--border);display:flex;align-items:center;justify-content:space-between;flex-shrink:0;}
  .dep-header h4{font-size:12px;font-weight:600;}
  .dep-head-main{display:flex;align-items:center;gap:8px;min-width:0;}
  .dep-head-badge{display:inline-flex;align-items:center;gap:4px;background:#2b341a;border:1px solid #4d7c0f;color:#d9f99d;border-radius:999px;padding:2px 8px;font-size:10px;font-weight:700;white-space:nowrap;}
  .dep-close{cursor:pointer;color:var(--text2);font-size:16px;line-height:1;padding:0 4px;}
  .dep-close:hover{color:var(--text);}
  .dep-body{flex:1;overflow-y:auto;padding:14px;}
  .dep-flow{display:flex;align-items:flex-start;gap:0;margin-bottom:16px;}
  .dep-col{display:flex;flex-direction:column;gap:6px;min-width:110px;} .dep-col.center{min-width:130px;}
  .dep-arrows{display:flex;flex-direction:column;justify-content:center;align-items:center;padding:0 6px;gap:6px;}
  .dep-box{border-radius:6px;padding:6px 8px;font-size:11px;line-height:1.3;}
  .dep-box.pred{background:#effaf4;border:1px solid #bfe7d1;color:#22684c;} .dep-box.pred.slipped{background:#fff1f4;border-color:#f0b8c3;color:#b64960;}
  .dep-box.focal{background:#eef2ff;border:2px solid var(--blue);color:var(--text);font-weight:600;} .dep-box.focal.slipped{background:#fff7ed;border-color:#f59e0b;color:#b45309;}
  .dep-box.suc{background:#eff6ff;border:1px solid #bfdbfe;color:#285ca6;} .dep-box.suc.slipped{background:#fff1f4;border-color:#f0b8c3;color:#b64960;}
  .dep-box .box-name{font-weight:500;margin-bottom:2px;} .dep-box .box-meta{font-size:10px;color:var(--text2);} .dep-box .box-slip{font-size:10px;color:#ffb74d;margin-top:2px;}
  .arr{font-size:18px;color:var(--text2);line-height:1;}
  .dep-section{margin-bottom:14px;} .dep-section h5{font-size:10px;text-transform:uppercase;letter-spacing:.5px;color:var(--text2);margin-bottom:6px;}
  .dep-detail-row{background:var(--surface2);border-radius:6px;padding:7px 10px;margin-bottom:5px;font-size:12px;}
  .dep-detail-row .dr-label{color:var(--text2);font-size:10px;margin-bottom:2px;}
  /* gantt */
  .gantt-outer{flex:1;display:flex;flex-direction:column;padding:0 18px 10px;overflow:hidden;}
  .gantt-controls{display:flex;align-items:center;gap:10px;padding:8px 0 6px;flex-shrink:0;flex-wrap:wrap;}
  .gantt-controls select{background:var(--surface);color:var(--text);border:1px solid var(--border);border-radius:8px;padding:4px 8px;font-size:12px;}
  .gantt-controls label{font-size:12px;color:var(--text2);display:flex;align-items:center;gap:5px;cursor:pointer;}
  .gantt-controls input[type=checkbox]{accent-color:var(--blue);}
  .gantt-legend{display:flex;gap:12px;font-size:11px;color:var(--text2);margin-left:auto;flex-wrap:wrap;}
  .gantt-legend span{display:flex;align-items:center;gap:4px;}
  .gantt-count{font-size:11px;color:var(--text2);font-weight:600;}
  .gantt-body{flex:1;display:flex;overflow:hidden;border:1px solid var(--border);border-radius:12px;background:var(--surface);box-shadow:var(--shadow);}
  .gantt-plot{flex:1;min-height:640px;background:var(--surface);}
  .gantt-scale-note{font-size:11px;color:var(--text2);font-style:italic;}
  #ganttTip{position:fixed;background:#ffffff;border:1px solid #cbd5e1;border-radius:6px;padding:8px 12px;font-size:11px;pointer-events:none;display:none;z-index:9999;max-width:360px;line-height:1.6;white-space:pre-wrap;color:#0f172a;box-shadow:0 6px 20px rgba(0,0,0,.15);}
  .loading-overlay{position:fixed;inset:0;background:rgba(244,247,251,.96);display:flex;align-items:center;justify-content:center;z-index:10000;transition:opacity .25s ease;}
  .loading-overlay.hide{opacity:0;pointer-events:none;}
  .loading-card{background:var(--surface);border:1px solid var(--border);border-radius:14px;padding:20px 24px;min-width:340px;box-shadow:var(--shadow);text-align:center;}
  .loading-title{font-size:14px;font-weight:700;color:var(--text);margin-bottom:12px;}
  .loading-stages{list-style:none;margin:0 0 14px;padding:0;text-align:left;}
  .loading-stages li{display:flex;align-items:center;gap:8px;font-size:12px;color:var(--text2);padding:3px 0;transition:color .2s;}
  .loading-stages li.active{color:var(--blue);font-weight:600;}
  .loading-stages li.done{color:#2f8f6b;}
  .loading-stages li .stage-icon{width:16px;text-align:center;flex-shrink:0;}
  .loading-bar-wrap{height:6px;background:#dde5ef;border-radius:4px;margin-bottom:8px;overflow:hidden;}
  .loading-bar{height:100%;border-radius:4px;background:var(--blue);transition:width .3s ease;width:0%;}
  .loading-pct{font-size:11px;color:var(--text2);}
  .loading-spin-sm{width:14px;height:14px;border:2px solid #d8e2ee;border-top-color:var(--blue);border-radius:50%;display:inline-block;animation:spin .85s linear infinite;}
  /* ── Dependency Board ── */
  .board-layout{display:grid;grid-template-columns:400px 1fr;gap:12px;flex:1;min-height:0;}
  .board-left{display:flex;flex-direction:column;min-height:0;gap:10px;}
  /* issue list */
  .board-list{overflow:auto;max-height:220px;padding:4px;min-height:0;overscroll-behavior:contain;scrollbar-gutter:stable both-edges;}
  .board-item{padding:9px 11px;border:1.5px solid var(--border);border-radius:10px;background:var(--surface);margin-bottom:5px;cursor:pointer;transition:border-color .15s,box-shadow .15s;}
  .board-item:hover{border-color:#93c5fd;box-shadow:0 1px 6px #3b82f615;}
  .board-item.active{border-color:var(--blue);background:#eef2ff;box-shadow:0 0 0 2px #c7d2fe;}
  .board-item-name{font-size:12.5px;font-weight:600;color:var(--text1);}
  .board-item-sub{font-size:11px;color:var(--text2);margin-top:2px;display:flex;gap:8px;flex-wrap:wrap;align-items:center;}
  .board-item-chip{display:inline-block;padding:1px 7px;border-radius:999px;font-size:10px;font-weight:600;background:#f1f5f9;border:1px solid var(--border);color:var(--text2);}
  .board-item-chip.red{background:#fff1f2;border-color:#fecdd3;color:#be123c;}
  .board-item-chip.amber{background:#fffbeb;border-color:#fde68a;color:#92400e;}
  /* root hero card */
  .board-root-card{background:linear-gradient(135deg,#eef2ff 0%,#e0e7ff 100%);border:1.5px solid #c7d2fe;border-radius:12px;padding:12px 14px;margin-bottom:2px;}
  .board-root-name{font-size:13px;font-weight:700;color:#1e1b4b;margin-bottom:6px;}
  .board-root-chips{display:flex;gap:6px;flex-wrap:wrap;}
  .board-root-chip{display:inline-flex;align-items:center;gap:4px;padding:2px 9px;border-radius:999px;font-size:10.5px;font-weight:600;border:1px solid;}
  .board-root-chip.owner{background:#ede9fe;border-color:#c4b5fd;color:#5b21b6;}
  .board-root-chip.dates{background:#fff7ed;border-color:#fed7aa;color:#c2410c;}
  .board-root-chip.pred{background:#dcfce7;border-color:#86efac;color:#166534;}
  .board-root-chip.suc{background:#dbeafe;border-color:#93c5fd;color:#1e40af;}
  .board-root-chip.pct{background:#f0fdf4;border-color:#86efac;color:#15803d;}
  /* section headers */
  .board-section-hdr{display:flex;align-items:center;gap:8px;margin:10px 0 4px;}
  .board-section-label{font-size:10.5px;font-weight:700;text-transform:uppercase;letter-spacing:.05em;color:var(--text2);}
  .board-section-count{display:inline-flex;align-items:center;justify-content:center;min-width:18px;height:18px;padding:0 5px;border-radius:999px;font-size:10px;font-weight:700;}
  .board-section-count.pred{background:#dcfce7;color:#166534;}
  .board-section-count.suc{background:#dbeafe;color:#1e40af;}
  /* tree */
  .board-tree{overflow:auto;flex:1;min-height:0;padding:4px 2px;overscroll-behavior:contain;scrollbar-gutter:stable both-edges;}
  .board-tree-inner{min-width:920px;padding-right:4px;}
  .board-node{position:relative;padding:7px 10px 7px 10px;border-radius:8px;margin:3px 0;border:1.5px solid transparent;transition:border-color .12s;}
  .board-node.branch::before{content:"";position:absolute;left:-1px;top:-4px;bottom:-4px;border-left:2px solid #e2e8f0;}
  .board-node.branch::after{content:"";position:absolute;left:-1px;top:50%;width:8px;border-top:2px solid #e2e8f0;}
  .board-node.pred{background:#f0fdf4;border-color:#bbf7d0;}
  .board-node.pred:hover{border-color:#4ade80;}
  .board-node.suc{background:#eff6ff;border-color:#bfdbfe;}
  .board-node.suc:hover{border-color:#60a5fa;}
  .board-node.active{box-shadow:0 0 0 2px #c7d2fe;}
  .board-node.driving{outline:2px solid #f59e0b;outline-offset:1px;}
  .board-node-row{display:flex;align-items:flex-start;gap:8px;}
  .board-node-toggle{flex-shrink:0;border:1.5px solid var(--border);background:var(--surface);border-radius:5px;cursor:pointer;width:20px;height:20px;line-height:1;font-size:12px;display:flex;align-items:center;justify-content:center;margin-top:1px;}
  .board-node-body{flex:1;min-width:0;cursor:pointer;}
  .board-node-name{font-size:12px;font-weight:600;color:var(--text1);white-space:nowrap;overflow:hidden;text-overflow:ellipsis;display:flex;align-items:center;gap:6px;}
  .board-dir{font-size:11px;color:#64748b;display:inline-flex;align-items:center;justify-content:center;border:1px solid #cbd5e1;border-radius:999px;padding:0 5px;background:#fff;}
  .board-node-meta{display:grid;grid-template-columns:110px 110px 160px 70px 90px 1fr;gap:6px;margin-top:5px;align-items:center;white-space:nowrap;}
  .board-col{font-size:10px;border:1px solid #e2e8f0;border-radius:6px;padding:2px 6px;background:#f8fafc;color:#475569;overflow:hidden;text-overflow:ellipsis;}
  .board-col.flags{display:flex;align-items:center;gap:4px;overflow:auto;padding:2px 4px;background:#fff;}
  .board-chip{display:inline-block;padding:1px 6px;border-radius:999px;font-size:10px;font-weight:600;border:1px solid;}
  .board-chip.date{background:#fff7ed;border-color:#fed7aa;color:#c2410c;}
  .board-chip.pct{background:#f0fdf4;border-color:#86efac;color:#15803d;}
  .board-chip.milestone{background:#faf5ff;border-color:#e9d5ff;color:#7c3aed;}
  .board-chip.driving{background:#fffbeb;border-color:#fcd34d;color:#92400e;}
  .board-chip.dep{background:#f8fafc;border-color:#cbd5e1;color:#475569;}
  .board-chip.status{font-size:9.5px;}
  .board-chip.status.danger{background:#fff1f2;border-color:#fecdd3;color:#be123c;}
  .board-chip.status.warn{background:#fffbeb;border-color:#fde68a;color:#92400e;}
  .board-chip.status.info{background:#eff6ff;border-color:#bfdbfe;color:#1d4ed8;}
  .board-chip.status.ok{background:#ecfdf5;border-color:#86efac;color:#166534;}
  /* toolbar */
  .board-toolbar{display:flex;align-items:center;gap:5px;flex-wrap:nowrap;}
  .board-btn{padding:3px 9px;border-radius:6px;border:1.5px solid var(--border);background:var(--surface);font-size:10.5px;font-weight:600;cursor:pointer;color:var(--text2);transition:all .12s;white-space:nowrap;}
  .board-btn:hover{border-color:#93c5fd;color:#1d4ed8;background:#eff6ff;}
  .board-btn.active{background:#1d4ed8;border-color:#1d4ed8;color:#fff;}
  .board-btn.amber.active{background:#d97706;border-color:#d97706;color:#fff;}
  .board-divider{width:1px;height:16px;background:var(--border);margin:0 1px;}
  /* gantt panel */
  .board-gantt-panel{display:flex;flex-direction:column;min-height:0;}
  .board-dep-detail{margin-top:8px;border:1px solid var(--border);border-radius:10px;background:#f8fafc;padding:8px;display:grid;grid-template-columns:1fr 1fr;gap:8px;min-height:86px;}
  .board-dep-col{background:#fff;border:1px solid #e2e8f0;border-radius:8px;padding:7px 8px;min-height:68px;}
  .board-dep-hd{font-size:10px;font-weight:700;text-transform:uppercase;letter-spacing:.04em;color:#64748b;margin-bottom:4px;}
  .board-dep-txt{font-size:11px;line-height:1.4;color:#0f172a;}
  .board-dep-meta{font-size:10px;color:#64748b;margin-top:4px;}

  /* actions board */
  .actions-layout{height:100%;display:flex;flex-direction:column;gap:8px;min-height:0;}
  .actions-toolbar{display:flex;align-items:center;justify-content:space-between;gap:10px;flex-wrap:wrap;}
  .actions-sub{font-size:11px;color:var(--text2);}
  .actions-board{flex:1;min-height:0;display:grid;grid-template-columns:repeat(3,minmax(280px,1fr));gap:10px;overflow-x:auto;padding-bottom:2px;}
  .actions-col{border:1px solid var(--border);border-radius:12px;background:var(--surface2);display:flex;flex-direction:column;min-height:0;overflow:hidden;}
  .actions-col-hd{padding:8px 10px;border-bottom:1px solid var(--border);display:flex;align-items:center;justify-content:space-between;font-size:11px;font-weight:700;text-transform:uppercase;letter-spacing:.04em;color:var(--text2);}
  .actions-col-body{padding:8px;display:flex;flex-direction:column;gap:8px;overflow:auto;min-height:0;overscroll-behavior:contain;scrollbar-gutter:stable both-edges;}
  .actions-col-immediate .actions-col-hd{background:#fff1f2;color:#be123c;}
  .actions-col-planned .actions-col-hd{background:#eff6ff;color:#1d4ed8;}
  .actions-col-monitor .actions-col-hd{background:#ecfdf5;color:#166534;}
  .action-card{background:var(--surface);border:1.5px solid var(--border);border-radius:10px;padding:9px 10px;display:flex;flex-direction:column;gap:6px;box-shadow:0 1px 3px rgba(15,23,42,.04);}
  .action-top{display:flex;flex-wrap:wrap;gap:5px;}
  .action-pill{display:inline-flex;align-items:center;gap:4px;padding:1px 7px;border-radius:999px;font-size:10px;font-weight:700;border:1px solid;max-width:100%;}
  .action-pill.ws{background:#eef2ff;border-color:#bfdbfe;color:#1e40af;}
  .action-pill.type{background:#fdf2f8;border-color:#fbcfe8;color:#9d174d;}
  .action-pill.ms{background:#fff7ed;border-color:#fed7aa;color:#9a3412;}
  .action-title{font-size:12px;font-weight:700;color:var(--text);line-height:1.35;}
  .action-linked{font-size:11px;color:var(--text2);}
  .action-desc{font-size:11px;color:#334155;line-height:1.45;}
  .action-chips{display:flex;flex-wrap:wrap;gap:5px;}
  .action-chip{display:inline-flex;align-items:center;padding:1px 7px;border-radius:999px;border:1px solid var(--border);background:#f8fafc;font-size:10px;font-weight:600;color:#475569;}
  .action-chip.owner{background:#ede9fe;border-color:#c4b5fd;color:#5b21b6;}
  .action-chip.date{background:#fff7ed;border-color:#fed7aa;color:#c2410c;}
  .action-chip.dep{background:#eef2ff;border-color:#bfdbfe;color:#1e40af;}
  .action-steps{margin:0;padding-left:16px;font-size:11px;color:#0f172a;line-height:1.45;}
  .action-steps li{margin:0 0 2px 0;}
  .action-impact{font-size:10px;color:var(--text2);padding-top:4px;border-top:1px dashed var(--border);}
  @keyframes spin{to{transform:rotate(360deg)}}
  ::-webkit-scrollbar{width:5px;height:5px;} ::-webkit-scrollbar-track{background:transparent;} ::-webkit-scrollbar-thumb{background:var(--border);border-radius:3px;}
</style>
</head>
<body>
<div id="loadingOverlay" class="loading-overlay" aria-live="polite">
  <div class="loading-card">
    <div class="loading-title">⏳ Preparing dashboard…</div>
    <ul class="loading-stages" id="loadingStages">
      <li id="ls0"><span class="stage-icon">⏳</span> Indexing task data</li>
      <li id="ls1"><span class="stage-icon">⏳</span> Rendering KPIs and status chart</li>
      <li id="ls2"><span class="stage-icon">⏳</span> Building milestones view</li>
      <li id="ls3"><span class="stage-icon">⏳</span> Populating risk / issue tiles</li>
      <li id="ls4"><span class="stage-icon">⏳</span> Rendering task table</li>
    </ul>
    <div class="loading-bar-wrap"><div class="loading-bar" id="loadingBar"></div></div>
    <div class="loading-pct" id="loadingPct">0%</div>
  </div>
</div>
<header>
  <h1>📊 Ametek SAP S4 — Schedule Dashboard</h1>
  <span id="healthBadge" class="health-badge">Loading…</span>
  <span class="gen" id="genTime"></span>
</header>
<div class="kpi-strip" id="kpiStrip"></div>
<div class="tab-bar">
  <button class="tab-btn active" id="tabTasks" onclick="switchTab('tasks')">📋 Tasks &amp; Issues</button>
  <button class="tab-btn" id="tabGantt" onclick="switchTab('gantt')">📅 Gantt / Milestones</button>
  <button class="tab-btn" id="tabActions" onclick="switchTab('actions')">🧠 Proposed Actions</button>
  <button class="tab-btn" id="tabBoard" onclick="switchTab('board')">🧩 Dependency Board</button>
  <button class="tab-btn" id="tabPlan" onclick="switchTab('plan')">🗂️ Plan Explorer</button>
</div>

<!-- TASKS TAB -->
<div id="tasksContent" style="display:contents">
  <div class="charts-strip">
    <div class="panel">
      <div class="panel-title">Tasks by Status</div>
      <div class="chart-wrap"><canvas id="statusChart"></canvas></div>
    </div>
    <div class="panel">
      <div class="panel-title">Milestone outlook — issue-linked milestones flagged</div>
      <div class="chart-wrap"><div id="wsChart" class="mini-milestone"></div></div>
    </div>
    <div class="panel">
      <div class="panel-title">Risks &amp; Issues by Type — click to filter</div>
      <div class="type-list" id="typeList"></div>
    </div>
  </div>
  <div class="main-area" id="mainArea">
    <div class="task-section">
      <div class="section-bar">
        <h3>Tasks</h3>
        <span class="filter-info" id="filterInfo"></span>
        <span class="filter-clear" id="filterClear" style="display:none" onclick="clearFilter()">✕ Clear filter</span>
        <span style="margin-left:auto;font-size:11px;color:var(--text2)">Click row → dep chain →</span>
      </div>
      <div class="tbl-wrap">
        <table>
          <thead><tr>
            <th>ID</th><th>Task Name</th><th>Milestone</th><th>Status</th><th>Start</th><th>Finish</th>
            <th>Baseline</th><th>Slip</th><th>↑Pred</th><th>↓Suc</th><th>Linked Issues / Risks</th>
          </tr></thead>
          <tbody id="taskBody"></tbody>
        </table>
      </div>
    </div>
    <div class="dep-panel" id="depPanel">
      <div class="dep-header">
        <div class="dep-head-main">
          <h4 id="depTitle">Dependency Chain</h4>
          <span id="depMeta"></span>
        </div>
        <span class="dep-close" onclick="closeDepPanel()">✕</span>
      </div>
      <div class="dep-body" id="depBody"></div>
    </div>
  </div>
</div>

<!-- ACTIONS TAB -->
<div id="actionsContent" style="display:none;flex:1;overflow:hidden;padding:10px 18px 12px;">
  <div class="panel" style="height:100%;overflow:hidden;">
    <div class="actions-layout">
      <div class="actions-toolbar">
        <div class="panel-title" style="margin-bottom:0;">Resolution Board — prioritized issue cards with linked task context</div>
        <div class="actions-sub" id="actionsSub"></div>
      </div>
      <div id="actionsBoard" class="actions-board"></div>
    </div>
  </div>
</div>

<!-- DEPENDENCY BOARD TAB -->
<div id="boardContent" style="display:none;flex:1;overflow:hidden;padding:10px 18px 12px;">
  <div class="board-layout">
    <!-- LEFT COLUMN -->
    <div class="board-left">
      <!-- Issue picker -->
      <div class="panel" style="padding:10px 10px 8px;">
        <div class="panel-title">Issues &amp; Risks</div>
        <div id="boardList" class="board-list"></div>
      </div>
      <!-- Tree panel -->
      <div class="panel" style="min-height:0;display:flex;flex-direction:column;padding:10px;">
        <div style="margin-bottom:8px;">
          <div style="display:flex;align-items:center;justify-content:space-between;">
            <div class="panel-title" style="margin-bottom:0;">Dependency Chain</div>
            <div class="board-toolbar">
              <button id="btnCritical" class="board-btn" onclick="toggleBoardMode('critical')">🎯 Critical</button>
              <button id="btnDriving" class="board-btn amber" onclick="toggleBoardMode('driving')">⚡ Driving</button>
              <div class="board-divider"></div>
              <button id="boardExpandAll" class="board-btn">+ All</button>
              <button id="boardCollapseAll" class="board-btn">− All</button>
            </div>
          </div>
          <div style="font-size:10px;color:var(--text2);margin-top:6px;display:flex;gap:8px;flex-wrap:wrap;">
            <span class="board-chip status danger" title="Task finish date is in the past and not complete">⛔ Past due</span>
            <span class="board-chip status warn" title="Task is currently slipping against baseline">🔻 Slipping</span>
            <span class="board-chip status info" title="Task start date is overdue">⏰ Overdue start</span>
            <span class="board-chip status ok" title="Task finishes within the next 14 days">⏳ Due ≤14d</span>
            <span style="margin-left:auto">↔ Scroll left/right for full columns</span>
          </div>
        </div>
        <div id="boardRootCard"></div>
        <div id="boardTree" class="board-tree"></div>
        <div id="boardDepDetail" class="board-dep-detail"></div>
      </div>
    </div>
    <!-- RIGHT COLUMN: Gantt -->
    <div class="panel board-gantt-panel">
      <div style="display:flex;align-items:center;justify-content:space-between;margin-bottom:6px;">
        <div class="panel-title" style="margin-bottom:0;">Schedule Chain — Gantt</div>
        <span id="boardGanttLegend" style="display:flex;gap:10px;font-size:10.5px;color:var(--text2);align-items:center;">
          <span><svg width="10" height="10"><rect width="10" height="10" rx="2" fill="#4b6bfb"/></svg> Root</span>
          <span><svg width="10" height="10"><rect width="10" height="10" rx="2" fill="#16a34a"/></svg> Predecessor</span>
          <span><svg width="10" height="10"><rect width="10" height="10" rx="2" fill="#2563eb"/></svg> Successor</span>
          <span><svg width="10" height="10"><polygon points="5,0 10,5 5,10 0,5" fill="#7c3aed"/></svg> Milestone</span>
          <span style="color:#d97706;">▽ Outline dependency</span>
        </span>
      </div>
      <div id="boardGantt" style="flex:1;min-height:0;"></div>
    </div>
  </div>
</div>

<!-- GANTT TAB -->
<div id="ganttContent" style="display:none;flex:1;flex-direction:column;overflow:hidden;">
  <div class="gantt-outer">
    <div class="gantt-controls">
      <label>Workstream:
        <select id="ganttWs" onchange="ganttRendered=false;renderGantt()">
          <option value="">All workstreams</option>
          ${wsOptions}
        </select>
      </label>
      <label>Timeline:
        <select id="ganttWindow" onchange="ganttRendered=false;renderGantt()">
          <option value="auto">Auto</option>
          <option value="3m">Next 3 months</option>
          <option value="6m">Next 6 months</option>
          <option value="12m">Next 12 months</option>
        </select>
      </label>
      <label>Scale:
        <select id="ganttScale" onchange="ganttRendered=false;renderGantt()">
          <option value="days">Days</option>
          <option value="weeks">Weeks</option>
          <option value="months" selected>Months</option>
        </select>
      </label>
      <label>Risk type:
        <select id="ganttRiskType" onchange="ganttRendered=false;renderGantt()">
          <option value="">All tasks</option>
          <option value="__any__">Any risk-linked task</option>
          ${riskTypeOptions}
        </select>
      </label>
      <label><input type="checkbox" id="ganttMilestones" checked onchange="ganttRendered=false;renderGantt()"> Milestone lane</label>
      <label><input type="checkbox" id="ganttIssues" checked onchange="ganttRendered=false;renderGantt()"> Has issues</label>
      <label><input type="checkbox" id="ganttSlipped" checked onchange="ganttRendered=false;renderGantt()"> Slipped</label>
      <span class="gantt-count" id="ganttCount"></span>
      <span class="gantt-scale-note" id="ganttScaleNote"></span>
      <div class="gantt-legend">
        <span><svg width="12" height="10"><polygon points="6,0 12,5 6,10 0,5" fill="#16a34a"/></svg> Milestone</span>
        <span><svg width="12" height="10"><rect width="12" height="10" rx="2" fill="#ea580c"/></svg> Slipped task</span>
        <span><svg width="12" height="10"><rect width="12" height="10" rx="2" fill="#2563eb"/></svg> Active task</span>
        <span><svg width="10" height="10"><circle cx="5" cy="5" r="5" fill="#dc2626"/></svg> Risk marker</span>
        <span style="font-weight:600">↑ / ↓ = shown predecessors/successors</span>
      </div>
    </div>
    <div class="gantt-body">
      <div class="gantt-plot" id="ganttPlot"></div>
    </div>
  </div>
</div>

<!-- PLAN EXPLORER TAB -->
<div id="planContent" style="display:none;flex:1;flex-direction:column;overflow:hidden;padding:10px 18px 12px;gap:10px;">
  <div class="panel" style="padding:8px 10px;flex-shrink:0;">
    <div class="panel-title">Rolling 4-week calendar (L3/L4 focus through mock data loads level)</div>
    <div id="planCalendar" style="min-height:90px;font-size:12px;color:var(--text2);">Loading…</div>
  </div>
  <div style="display:grid;grid-template-columns:minmax(420px,1.1fr) minmax(420px,1fr);gap:10px;flex:1;min-height:0;">
    <div class="panel" style="min-height:0;display:flex;flex-direction:column;">
      <div style="display:flex;align-items:center;gap:8px;margin-bottom:6px;">
        <div class="panel-title" style="margin-bottom:0;">Whole Plan Hierarchy</div>
        <button id="planExpandAll" style="font-size:11px;padding:3px 8px;border:1px solid var(--border);background:var(--surface2);border-radius:6px;cursor:pointer;">Expand all</button>
        <button id="planCollapseAll" style="font-size:11px;padding:3px 8px;border:1px solid var(--border);background:var(--surface2);border-radius:6px;cursor:pointer;">Collapse all</button>
        <span id="planCount" style="margin-left:auto;font-size:11px;color:var(--text2);"></span>
      </div>
      <div class="tbl-wrap" style="box-shadow:none;border-radius:10px;">
        <table>
          <thead>
            <tr>
              <th>Task</th><th>Lvl</th><th>Owner</th><th>Finish</th><th>Indicators</th>
            </tr>
          </thead>
          <tbody id="planBody"></tbody>
        </table>
      </div>
    </div>
    <div class="panel" style="min-height:0;display:flex;flex-direction:column;">
      <div class="panel-title">Adjacent plan gantt (visible hierarchy rows)</div>
      <div id="planGantt" style="flex:1;min-height:0;"></div>
    </div>
  </div>
</div>
<div id="ganttTip"></div>

<script>
const DATA = ${data};
let activeFilter = null;
let selectedTaskId = null;
let ganttRendered = false;
let planRendered = false;
let boardRendered = false;
const taskById = new Map(DATA.tasks.map(t => [t.id, t]));
const taskByUid = new Map(DATA.tasks.filter(t => t.uid != null).map(t => [t.uid, t]));
let planState = null;
let planGanttTimer = null;
const boardState = { selectedActionId: null, selectedNodeId: null, expanded: new Set(), tree: null, nodeMap: new Map(), rootTask: null, criticalOnly: false, drivingOnly: false };
const fmt = d => d ? new Date(d).toLocaleDateString("en-US",{month:"short",day:"numeric",year:"2-digit"}) : "—";
const fmtShort = d => d ? new Date(d).toLocaleDateString("en-US",{month:"short",day:"numeric"}) : "—";
const esc = s => String(s||"").replace(/&/g,"&amp;").replace(/</g,"&lt;").replace(/>/g,"&gt;");
const trunc = (s,n) => s && s.length>n ? s.slice(0,n)+"…" : (s||"");
const riskKey = r => (r.type || "Risk") + " — " + (r.category || "(Uncategorized)");
const topRiskKey = task => task.linkedRisks.length ? riskKey(task.linkedRisks[0]) : "No linked risk";
const PERF = { initialTaskRows: 180, maxTaskRows: 320 };
let atRiskTasksCache = null;

function scheduleNonBlocking(fn) {
  if (typeof requestIdleCallback === 'function') {
    requestIdleCallback(() => fn(), { timeout: 150 });
  } else {
    setTimeout(fn, 0);
  }
}

function setLoadingText(msg) {
  const el = document.getElementById('loadingSub');
  if (el && msg) el.textContent = msg;
}

function hideLoading() {
  const overlay = document.getElementById('loadingOverlay');
  if (!overlay) return;
  overlay.classList.add('hide');
  setTimeout(() => overlay.remove(), 280);
}

function buildPlanState() {
  const byId = new Map(DATA.tasks.map(t => [t.id, {
    id: t.id,
    uid: t.uid,
    name: t.name,
    assignedTo: t.assignedTo,
    start: t.start,
    finish: t.finish,
    outlineLevel: t.outlineLevel || null,
    parentUid: t.parentUid || null,
    workstream: t.workstream,
    task: t,
    children: [],
    parentId: null
  }]));
  const nodes = [...byId.values()];

  nodes.forEach(n => {
    if (n.parentUid != null && taskByUid.has(n.parentUid)) {
      const parentTask = taskByUid.get(n.parentUid);
      const parent = byId.get(parentTask.id);
      if (parent && parent.id !== n.id) {
        n.parentId = parent.id;
        parent.children.push(n.id);
      }
    }
  });

  const roots = nodes.filter(n => !n.parentId).sort((a, b) => {
    if (a.outlineLevel != null && b.outlineLevel != null && a.outlineLevel !== b.outlineLevel) return a.outlineLevel - b.outlineLevel;
    return (a.uid || 999999) - (b.uid || 999999) || a.name.localeCompare(b.name);
  }).map(n => n.id);

  nodes.forEach(n => n.children.sort((aId, bId) => {
    const a = byId.get(aId), b = byId.get(bId);
    return (a.uid || 999999) - (b.uid || 999999) || a.name.localeCompare(b.name);
  }));

  const expanded = new Set();
  nodes.forEach(n => {
    if (n.children.length && (n.outlineLevel == null || n.outlineLevel <= 2)) expanded.add(n.id);
  });

  const issueRollup = new Map();
  const mark = (id) => {
    if (issueRollup.has(id)) return issueRollup.get(id);
    const node = byId.get(id);
    if (!node) return false;
    const selfIssue = (node.task.linkedRisks || []).some(r => /issue/i.test(r.type));
    if (selfIssue) {
      issueRollup.set(id, true);
      return true;
    }
    for (const cid of node.children) {
      if (mark(cid)) {
        issueRollup.set(id, true);
        return true;
      }
    }
    issueRollup.set(id, false);
    return false;
  };
  roots.forEach(id => mark(id));

  return { byId, roots, expanded, issueRollup };
}

function visiblePlanRows() {
  if (!planState) planState = buildPlanState();
  const rows = [];
  const walk = (id, depth) => {
    const n = planState.byId.get(id);
    if (!n) return;
    rows.push({ node: n, depth });
    if (n.children.length && planState.expanded.has(n.id)) {
      n.children.forEach(cid => walk(cid, depth + 1));
    }
  };
  planState.roots.forEach(id => walk(id, 0));
  return rows;
}

function hasIssueRollup(nodeId) {
  if (!planState) return false;
  return !!planState.issueRollup.get(nodeId);
}

function renderPlanCalendar() {
  const host = document.getElementById('planCalendar');
  const now = new Date();
  const to = new Date(now.getTime() + 28 * 86400000);
  const focus = DATA.tasks
    .filter(t => (t.outlineLevel === 3 || t.outlineLevel === 4) && t.pct < 100 && (t.start || t.finish))
    .filter(t => {
      const d = new Date(t.start || t.finish);
      return d >= now && d <= to;
    })
    .sort((a, b) => (a.finish || a.start || '').localeCompare(b.finish || b.start || ''))
    .slice(0, 16);

  if (!focus.length) {
    host.innerHTML = '<div style="padding:8px 2px">No L3/L4 items in the next 4 weeks.</div>';
    return;
  }

  host.innerHTML = focus.map(t => {
    const flag = t.linkedRisks.length ? '⚠' : '•';
    const date = fmt(t.start || t.finish);
    return '<div style="padding:4px 0;border-bottom:1px dashed var(--border)">' +
      '<span style="display:inline-block;min-width:84px;color:var(--text2)">' + esc(date) + '</span>' +
      '<span style="font-weight:600;color:var(--text)">L' + esc(String(t.outlineLevel || '?')) + '</span> ' + flag + ' ' +
      esc(trunc(t.name, 70)) +
      '<span style="color:var(--text2)"> — ' + esc(t.assignedTo || 'Unassigned') + '</span>' +
      '</div>';
  }).join('');
}

function renderPlanExplorer() {
  if (!planState) planState = buildPlanState();
  const rows = visiblePlanRows();
  const body = document.getElementById('planBody');
  document.getElementById('planCount').textContent = rows.length + ' visible rows';

  if (!rows.length) {
    body.innerHTML = '<tr><td colspan="5" class="empty">No plan rows available.</td></tr>';
    return;
  }

  body.innerHTML = rows.slice(0, 1200).map(r => {
    const n = r.node;
    const canExpand = n.children.length > 0;
    const isOpen = canExpand && planState.expanded.has(n.id);
    const issueFlag = hasIssueRollup(n.id);
    const slipFlag = n.task.isSlipped;
    const ind = [
      issueFlag ? '<span class="pill issue">Issue</span>' : '',
      slipFlag ? '<span class="pill slip">Slip</span>' : '',
      n.task.milestone ? '<span class="pill done">🏁</span>' : ''
    ].filter(Boolean).join(' ');
    return '<tr>' +
      '<td>' +
        '<div style="display:flex;align-items:center;gap:6px;padding-left:' + (r.depth * 14) + 'px">' +
          (canExpand
            ? '<button onclick="togglePlanNode(\\\'' + n.id + '\\\')" style="border:1px solid var(--border);background:var(--surface2);border-radius:4px;width:18px;height:18px;cursor:pointer;font-size:11px;line-height:1">' + (isOpen ? '−' : '+') + '</button>'
            : '<span style="display:inline-block;width:18px"></span>') +
          '<span style="font-weight:' + (r.depth <= 1 ? 600 : 500) + '">' + esc(trunc(n.name, 70)) + '</span>' +
        '</div>' +
      '</td>' +
      '<td>' + esc(String(n.outlineLevel || '—')) + '</td>' +
      '<td>' + esc(trunc(n.assignedTo || 'Unassigned', 22)) + '</td>' +
      '<td>' + esc(fmt(n.finish)) + '</td>' +
      '<td>' + (ind || '<span style="color:var(--text2)">—</span>') + '</td>' +
      '</tr>';
  }).join('');

  if (planGanttTimer) clearTimeout(planGanttTimer);
  planGanttTimer = setTimeout(() => renderPlanGantt(rows), 0);
}

function togglePlanNode(id) {
  if (!planState) return;
  if (planState.expanded.has(id)) planState.expanded.delete(id);
  else planState.expanded.add(id);
  renderPlanExplorer();
}

function renderPlanGantt(rows) {
  const host = document.getElementById('planGantt');
  const taskRows = rows.filter(r => r.node.task.pct < 100 && (r.node.start || r.node.finish)).slice(0, 140);
  if (!taskRows.length) {
    host.innerHTML = '<div style="padding:10px;color:var(--text2)">No visible dated tasks to chart.</div>';
    return;
  }

  const minD = new Date(Math.min(...taskRows.map(r => new Date(r.node.start || r.node.finish).getTime())));
  const maxD = new Date(Math.max(...taskRows.map(r => new Date(r.node.finish || r.node.start).getTime())));
  minD.setDate(minD.getDate() - 3);
  maxD.setDate(maxD.getDate() + 14);

  const y = taskRows.map(r => ' '.repeat(Math.min(10, r.depth * 2)) + trunc(r.node.name, 40));
  const base = taskRows.map(r => r.node.start || r.node.finish);
  const dur = taskRows.map(r => Math.max(86400000, new Date(r.node.finish || r.node.start).getTime() - new Date(r.node.start || r.node.finish).getTime() + 86400000));

  Plotly.react(host, [{
    type: 'bar',
    orientation: 'h',
    y,
    base,
    x: dur,
    marker: {
      color: taskRows.map(r => r.node.task.isSlipped ? '#d97757' : hasIssueRollup(r.node.id) ? '#5b7cff' : '#94a3b8')
    },
    text: taskRows.map(r => fmtShort(r.node.start || r.node.finish) + ' → ' + fmtShort(r.node.finish || r.node.start)),
    textposition: 'inside',
    insidetextanchor: 'middle',
    textfont: { color: '#fff', size: 10 },
    hovertemplate: taskRows.map(r => '<b>' + esc(r.node.name) + '</b><br>Level: ' + esc(String(r.node.outlineLevel || '—')) + '<br>Start: ' + fmt(r.node.start) + '<br>Finish: ' + fmt(r.node.finish) + '<extra></extra>')
  }], {
    margin: { l: 240, r: 12, t: 10, b: 36 },
    paper_bgcolor: '#ffffff',
    plot_bgcolor: '#ffffff',
    barmode: 'overlay',
    height: Math.max(560, taskRows.length * 24),
    xaxis: {
      type: 'date',
      range: [minD.toISOString(), maxD.toISOString()],
      showgrid: true,
      gridcolor: '#e7edf4',
      tickfont: { color: '#5f7185', size: 10 }
    },
    yaxis: {
      automargin: true,
      autorange: 'reversed',
      tickfont: { color: '#334155', size: 10 }
    },
    showlegend: false
  }, { displayModeBar: false, responsive: true });
}

function renderMilestoneMini() {
  const host = document.getElementById("wsChart");
  if (!host) return;
  const now = new Date();
  const from = new Date(now.getTime() - 14 * 86400000);
  const to = new Date(now.getTime() + 14 * 86400000);
  const milestones = DATA.tasks
    .filter(t => t.milestone && t.pct < 100 && (t.finish || t.start))
    .filter(t => {
      const d = new Date(t.finish || t.start);
      return d >= from && d <= to;
    })
    .sort((a, b) => (a.finish || a.start || "").localeCompare(b.finish || b.start || ""))
    .slice(0, 12);

  if (!milestones.length) {
    host.innerHTML = '<div style="padding:12px;color:var(--text2);font-size:12px">No open milestones in the ±2 week window.</div>';
    return;
  }

  host.innerHTML = '<div style="display:flex;flex-direction:column;gap:5px;padding:4px 2px">' +
    milestones.map(m => {
      const hasRisk = m.linkedRisks.length > 0;
      const icon = hasRisk ? '⚠️' : '🏁';
      const color = hasRisk ? '#d85b72' : '#2f8f6b';
      const date = fmt(m.finish || m.start);
      return '<div style="display:flex;align-items:center;gap:8px;padding:4px 6px;border-radius:6px;background:var(--surface2);border-left:3px solid ' + color + '">' +
        '<span style="font-size:13px">' + icon + '</span>' +
        '<span style="font-size:11px;color:var(--text2);white-space:nowrap;min-width:70px">' + esc(date) + '</span>' +
        '<span style="font-size:12px;font-weight:600;color:var(--text);flex:1;overflow:hidden;text-overflow:ellipsis;white-space:nowrap">' + esc(trunc(m.name, 50)) + '</span>' +
        '<span style="font-size:11px;color:var(--text2);white-space:nowrap">' + esc(trunc(m.assignedTo || 'Unassigned', 18)) + '</span>' +
        '</div>';
    }).join('') +
  '</div>';
}

// ── Plotly lazy loader ─────────────────────────────────────────────────────
let _plotlyPromise = null;
function ensurePlotly() {
  if (window.Plotly) return Promise.resolve();
  if (_plotlyPromise) return _plotlyPromise;
  _plotlyPromise = new Promise((resolve, reject) => {
    const s = document.createElement('script');
    s.src = 'https://cdn.jsdelivr.net/npm/plotly.js-dist-min@2.35.2/plotly.min.js';
    s.onload = resolve;
    s.onerror = reject;
    document.head.appendChild(s);
  });
  return _plotlyPromise;
}

// ── Tab switching ──────────────────────────────────────────────────────────
function switchTab(tab) {
  const isGantt = tab === "gantt";
  const isActions = tab === "actions";
  const isBoard = tab === "board";
  const isPlan = tab === "plan";
  document.getElementById("tabTasks").classList.toggle("active", !isGantt && !isActions && !isPlan && !isBoard);
  document.getElementById("tabGantt").classList.toggle("active", isGantt);
  document.getElementById("tabActions").classList.toggle("active", isActions);
  document.getElementById("tabBoard").classList.toggle("active", isBoard);
  document.getElementById("tabPlan").classList.toggle("active", isPlan);
  const tc = document.getElementById("tasksContent");
  tc.style.display = isGantt || isActions || isPlan || isBoard ? "none" : "contents";
  const ac = document.getElementById("actionsContent");
  ac.style.display = isActions ? "flex" : "none";
  const bc = document.getElementById("boardContent");
  bc.style.display = isBoard ? "flex" : "none";
  const gc = document.getElementById("ganttContent");
  gc.style.display = isGantt ? "flex" : "none";
  const pc = document.getElementById("planContent");
  pc.style.display = isPlan ? "flex" : "none";
  if (isGantt && !ganttRendered) {
    gc.innerHTML = '<div style="padding:24px;color:var(--text2);font-size:13px">⏳ Loading Gantt library…</div>';
    ensurePlotly().then(() => { renderGantt(); ganttRendered = true; });
  }
  if (isActions) renderActions();
  if (isBoard && !boardRendered) {
    ensurePlotly().then(() => {
      renderBoardTab();
      boardRendered = true;
    });
  }
  if (isPlan && !planRendered) {
    renderPlanCalendar();
    ensurePlotly().then(() => {
      renderPlanExplorer();
      planRendered = true;
    });
  }
}

function findActionForTask(taskId) {
  if (!taskId) return null;
  return (DATA.actionItems || []).find(a => a.taskId === taskId) || null;
}

function taskSignalList(task, driving) {
  if (!task) return [];
  const now = Date.now();
  const finishMs = task.finish ? new Date(task.finish).getTime() : null;
  const out = [];
  if (finishMs && task.pct < 100 && finishMs < now) {
    out.push({ cls: 'danger', icon: '⛔', label: 'Past due', tip: 'Finish date is in the past and task is not complete.' });
  }
  if (task.isSlipped) {
    out.push({ cls: 'warn', icon: '🔻', label: 'Slipping', tip: 'Task is behind baseline schedule.' });
  }
  if (task.isOverdueStart) {
    out.push({ cls: 'info', icon: '⏰', label: 'Overdue start', tip: 'Planned start date is in the past and work has not started as expected.' });
  }
  if (task.isDue14) {
    out.push({ cls: 'ok', icon: '⏳', label: 'Due ≤14d', tip: 'Task is due within the next 14 days.' });
  }
  if (driving) {
    out.push({ cls: 'warn', icon: '⚡', label: 'Driving', tip: 'This predecessor directly drives the successor with near-zero float.' });
  }
  if (task.milestone) {
    out.push({ cls: 'info', icon: '🏁', label: 'Milestone', tip: 'Milestone-type schedule node.' });
  }
  return out;
}

function renderBoardDepDetail(task) {
  const host = document.getElementById('boardDepDetail');
  if (!host) return;
  if (!task) {
    host.innerHTML =
      '<div class="board-dep-col"><div class="board-dep-hd">Issue</div><div class="board-dep-txt">Click any dependency node to see its linked issue context.</div></div>' +
      '<div class="board-dep-col"><div class="board-dep-hd">Resolution</div><div class="board-dep-txt">Suggested resolution steps will appear here next to the issue.</div></div>';
    return;
  }
  const action = findActionForTask(task.id) || (boardState.selectedActionId ? (DATA.actionItems || []).find(a => a.id === boardState.selectedActionId) : null);
  const issue = action?.plainIssue || (task.linkedRisks?.[0]?.name || 'No linked issue text available for this dependency.');
  const impact = action?.plainImpact || action?.linkageImpact || 'No impact description provided.';
  const resolutionSteps = (action?.actions || []).slice(0, 4);
  const resolutionHtml = resolutionSteps.length
    ? '<ul class="action-steps" style="padding-left:15px">' + resolutionSteps.map(s => '<li>' + esc(s) + '</li>').join('') + '</ul>'
    : '<div class="board-dep-txt">No explicit resolution steps found; review predecessor/successor constraints and owner plan.</div>';

  host.innerHTML =
    '<div class="board-dep-col">' +
      '<div class="board-dep-hd">Issue</div>' +
      '<div class="board-dep-txt">' + esc(trunc(issue, 240)) + '</div>' +
      '<div class="board-dep-meta">Impact: ' + esc(trunc(impact, 220)) + '</div>' +
    '</div>' +
    '<div class="board-dep-col">' +
      '<div class="board-dep-hd">Resolution</div>' +
      resolutionHtml +
      '<div class="board-dep-meta">Owner: ' + esc(action?.owner || task.assignedTo || 'Unassigned') + ' · Milestone: ' + esc(fmt(action?.milestoneDate || task.finish)) + '</div>' +
    '</div>';
}

function selectBoardNode(nodeId) {
  boardState.selectedNodeId = nodeId;
  const node = boardState.nodeMap.get(nodeId);
  renderBoardDepDetail(node?.task || null);
  renderBoardTree();
}

function selectBoardRoot() {
  boardState.selectedNodeId = '__root__';
  renderBoardDepDetail(boardState.rootTask || null);
  renderBoardTree();
}

function buildBoardSubtree(task, relation, depth, maxDepth, trail, nodeMap, idRef, parentTaskId) {
  const node = { id: ++idRef.value, task, relation, depth, parentTaskId: parentTaskId || null, children: [] };
  nodeMap.set(node.id, node);
  if (depth >= maxDepth) return node;
  const nextIds = relation === 'pred' ? (task.predIds || []) : (task.sucIds || []);
  for (const nid of nextIds) {
    const childTask = taskById.get(nid);
    if (!childTask) continue;
    const key = relation + ':' + nid;
    if (trail.has(key)) continue;
    const nextTrail = new Set(trail);
    nextTrail.add(key);
    node.children.push(buildBoardSubtree(childTask, relation, depth + 1, maxDepth, nextTrail, nodeMap, idRef, task.id));
  }
  return node;
}

function criticalScore(task) {
  if (!task) return -1;
  let score = 0;
  if (task.isSlipped) score += 40;
  if (task.isOverdueStart) score += 20;
  if (task.isDue14) score += 10;
  if (task.slipDays) score += Math.min(30, task.slipDays);
  if (task.pct < 100) score += 8;
  score += (task.predCount || 0) * 2 + (task.sucCount || 0) * 2;
  return score;
}

function pickCriticalChild(tasks) {
  if (!tasks.length) return null;
  return [...tasks].sort((a, b) => {
    const s = criticalScore(b) - criticalScore(a);
    if (s !== 0) return s;
    return String(a.finish || a.start || '').localeCompare(String(b.finish || b.start || ''));
  })[0] || null;
}

function buildBoardTree(rootTask) {
  const nodeMap = new Map();
  const idRef = { value: 0 };
  let predCandidates = (rootTask.predIds || []).map(pid => {
    const t = taskById.get(pid);
    if (!t) return null;
    return t;
  }).filter(Boolean);
  let sucCandidates = (rootTask.sucIds || []).map(sid => {
    const t = taskById.get(sid);
    if (!t) return null;
    return t;
  }).filter(Boolean);

  if (boardState.criticalOnly) {
    const predPick = pickCriticalChild(predCandidates);
    const sucPick = pickCriticalChild(sucCandidates);
    predCandidates = predPick ? [predPick] : [];
    sucCandidates = sucPick ? [sucPick] : [];
  }
  if (boardState.drivingOnly) {
    predCandidates = predCandidates.filter(t => isDrivingPred(t, rootTask));
  }

  const predRoots = predCandidates.map(t =>
    buildBoardSubtree(t, 'pred', 1, 3, new Set(['pred:' + rootTask.id]), nodeMap, idRef, rootTask.id)
  );
  const sucRoots = sucCandidates.map(t =>
    buildBoardSubtree(t, 'suc', 1, 3, new Set(['suc:' + rootTask.id]), nodeMap, idRef, rootTask.id)
  );
  return { predRoots, sucRoots, nodeMap };
}

function isDrivingPred(predTask, childTask) {
  // A predecessor is "driving" when its finish is within 1 day of the child's start (no float)
  if (!predTask || !childTask) return false;
  const pf = predTask.finish ? new Date(predTask.finish).getTime() : null;
  const cs = childTask.start ? new Date(childTask.start).getTime() : null;
  if (!pf || !cs) return false;
  return (cs - pf) <= 86400000; // ≤ 1 day gap
}

function boardNodeRow(node) {
  const canExpand = node.children.length > 0;
  const isOpen = boardState.expanded.has(node.id);
  const isSelected = boardState.selectedNodeId === node.id;
  const cls = node.relation === 'pred' ? 'pred' : 'suc';
  const pad = node.depth * 18;
  const parentTask = node.relation === 'pred' && node.parentTaskId ? taskById.get(node.parentTaskId) : null;
  const driving = node.relation === 'pred' && isDrivingPred(node.task, parentTask || boardState.rootTask);
  const btn = canExpand
    ? '<button class="board-node-toggle" onclick="toggleBoardNode(' + node.id + ')">' + (isOpen ? '−' : '+') + '</button>'
    : '<span style="display:inline-block;width:20px;"></span>';
  const signalBadges = taskSignalList(node.task, driving)
    .map(s => '<span class="board-chip status ' + s.cls + '" title="' + esc(s.tip) + '">' + s.icon + ' ' + esc(s.label) + '</span>')
    .join('');
  const rowTitle = [
    'Task: ' + (node.task.name || '—'),
    'Owner: ' + (node.task.assignedTo || 'Unassigned'),
    'Start: ' + fmt(node.task.start),
    'Finish: ' + fmt(node.task.finish),
    'Progress: ' + (node.task.pct || 0) + '%',
    'Pred: ' + (node.task.predCount || 0) + ' · Suc: ' + (node.task.sucCount || 0)
  ].join('\n');
  const drivingCls = driving ? ' driving' : '';
  const activeCls = isSelected ? ' active' : '';
  const row = '<div class="board-node branch ' + cls + drivingCls + activeCls + '" style="margin-left:' + pad + 'px">' +
    '<div class="board-node-row">' + btn +
    '<div class="board-node-body" onclick="selectBoardNode(' + node.id + ')" title="' + esc(rowTitle) + '">' +
    '<div class="board-node-name"><span class="board-dir" title="Dependency flow direction">↓</span>' + esc(trunc(node.task.name, 60)) + '</div>' +
    '<div class="board-node-meta">' +
      '<span class="board-col" title="Start date">Start: ' + esc(fmt(node.task.start)) + '</span>' +
      '<span class="board-col" title="Finish date">Finish: ' + esc(fmt(node.task.finish)) + '</span>' +
      '<span class="board-col" title="Assigned owner">Owner: ' + esc(trunc(node.task.assignedTo || 'Unassigned', 20)) + '</span>' +
      '<span class="board-col" title="Percent complete">%: ' + esc(String(node.task.pct || 0)) + '</span>' +
      '<span class="board-col" title="Dependency counts">↑' + esc(String(node.task.predCount || 0)) + ' ↓' + esc(String(node.task.sucCount || 0)) + '</span>' +
      '<span class="board-col flags">' + signalBadges + '</span>' +
    '</div>' +
    '</div></div></div>';
  if (!canExpand || !isOpen) return row;
  return row + node.children.map(boardNodeRow).join('');
}

function renderBoardTree() {
  const host = document.getElementById('boardTree');
  const rootCard = document.getElementById('boardRootCard');
  if (!boardState.rootTask || !boardState.tree) {
    if (rootCard) rootCard.innerHTML = '';
    renderBoardDepDetail(null);
    host.innerHTML = '<div style="color:var(--text2);padding:8px">Select an issue to explore its dependency chain.</div>';
    return;
  }
  const root = boardState.rootTask;
  const rootActive = boardState.selectedNodeId === '__root__' ? ' style="box-shadow:0 0 0 2px #c7d2fe;cursor:pointer"' : ' style="cursor:pointer"';
  // Root hero card
  if (rootCard) {
    const slipped = root.isSlipped ? '<span class="board-root-chip dates">🔴 Slipped ' + esc(String(root.slipDays || '')) + 'd</span>' : '';
    rootCard.innerHTML = '<div class="board-root-card" onclick="selectBoardRoot()"' + rootActive + '>' +
      '<div class="board-root-name">' + (root.milestone ? '🏁 ' : '') + esc(trunc(root.name, 70)) + '</div>' +
      '<div class="board-root-chips">' +
        (root.assignedTo ? '<span class="board-root-chip owner">👤 ' + esc(trunc(root.assignedTo, 24)) + '</span>' : '') +
        '<span class="board-root-chip dates">📅 ' + esc(fmt(root.start)) + ' → ' + esc(fmt(root.finish)) + '</span>' +
        '<span class="board-root-chip pct">' + esc(String(root.pct || 0)) + '% done</span>' +
        '<span class="board-root-chip pred">↑ ' + esc(String(root.predCount || 0)) + ' pred</span>' +
        '<span class="board-root-chip suc">↓ ' + esc(String(root.sucCount || 0)) + ' suc</span>' +
        slipped +
      '</div></div>';
  }
  const predRoots = boardState.tree.predRoots;
  const sucRoots = boardState.tree.sucRoots;
  const predHdr = '<div class="board-section-hdr"><span class="board-section-label">Predecessors</span><span class="board-section-count pred">' + predRoots.length + '</span></div>';
  const sucHdr = '<div class="board-section-hdr"><span class="board-section-label">Successors</span><span class="board-section-count suc">' + sucRoots.length + '</span></div>';
  const predBody = predRoots.length ? predRoots.map(boardNodeRow).join('') : '<div style="font-size:12px;color:var(--text2);padding:4px 6px">No predecessors</div>';
  const sucBody = sucRoots.length ? sucRoots.map(boardNodeRow).join('') : '<div style="font-size:12px;color:var(--text2);padding:4px 6px">No successors</div>';
  host.innerHTML = '<div class="board-tree-inner">' + predHdr + predBody + sucHdr + sucBody + '</div>';
}

function collectBoardTasks() {
  if (!boardState.rootTask) return [];
  const out = [{ task: boardState.rootTask, relation: 'ROOT', depth: 0, parentTaskId: null }];
  const seen = new Set([boardState.rootTask.id]);
  if (!boardState.tree) return out;
  for (const node of boardState.tree.nodeMap.values()) {
    if (node?.task && !seen.has(node.task.id)) {
      seen.add(node.task.id);
      out.push({
        task: node.task,
        relation: node.relation === 'pred' ? 'PRED' : 'SUC',
        depth: node.depth || 1,
        parentTaskId: node.parentTaskId || boardState.rootTask.id
      });
    }
  }
  return out;
}

function renderBoardGantt() {
  const host = document.getElementById('boardGantt');
  const nodes = collectBoardTasks().filter(x => x.task && (x.task.start || x.task.finish));
  if (!nodes.length) {
    host.innerHTML = '<div style="padding:12px;color:var(--text2)">No dated tasks to chart.</div>';
    return;
  }

  const relOrder = { ROOT: 0, PRED: 1, SUC: 2 };
  const rows = nodes.sort((a, b) => {
    if (relOrder[a.relation] !== relOrder[b.relation]) return relOrder[a.relation] - relOrder[b.relation];
    if ((a.depth || 0) !== (b.depth || 0)) return (a.depth || 0) - (b.depth || 0);
    return String(a.task.finish || a.task.start).localeCompare(String(b.task.finish || b.task.start));
  });

  const minD = new Date(Math.min(...rows.map(r => new Date(r.task.start || r.task.finish).getTime())));
  const maxD = new Date(Math.max(...rows.map(r => new Date(r.task.finish || r.task.start).getTime())));
  minD.setDate(minD.getDate() - 2);
  maxD.setDate(maxD.getDate() + 10);

  const y = rows.map(r => {
    const indent = '\u00a0'.repeat(Math.min(8, (r.depth || 0) * 3));
    const tag = r.relation === 'ROOT' ? '\u25cf' : (r.relation === 'PRED' ? '\u2191' : '\u2193');
    const typeTag = r.task.milestone ? '\ud83c\udfc1' : '';
    return indent + tag + ' ' + typeTag + trunc(r.task.name, 42);
  });

  const taskRows = rows.filter(r => !r.task.milestone);
  const milestoneRows = rows.filter(r => r.task.milestone);
  const yById = new Map(rows.map((r, i) => [r.task.id, y[i]]));

  const traces = [];
  if (taskRows.length) {
    traces.push({
      type: 'bar',
      orientation: 'h',
      y: taskRows.map(r => yById.get(r.task.id)),
      base: taskRows.map(r => r.task.start || r.task.finish),
      x: taskRows.map(r => Math.max(86400000, new Date(r.task.finish || r.task.start).getTime() - new Date(r.task.start || r.task.finish).getTime() + 86400000)),
      marker: {
        color: taskRows.map(r => r.relation === 'ROOT' ? '#4b6bfb' : r.relation === 'PRED' ? '#16a34a' : '#2563eb'),
        line: { width: 1, color: '#1e293b22' }
      },
      text: taskRows.map(r => fmtShort(r.task.start || r.task.finish) + ' → ' + fmtShort(r.task.finish || r.task.start)),
      textposition: 'auto',
      textfont: { color: '#fff', size: 10 },
      hovertemplate: taskRows.map(r => {
        const sig = taskSignalList(r.task, false).map(s => s.icon + ' ' + s.label).join(' · ') || 'None';
        return '<b>' + esc(r.task.name) + '</b><br>Relation: ' + r.relation + ' · Level ' + (r.depth || 0) + '<br>Type: Task<br>Start: ' + fmt(r.task.start) + '<br>Finish: ' + fmt(r.task.finish) + '<br>% Complete: ' + (r.task.pct || 0) + '<br>Owner: ' + esc(r.task.assignedTo || 'Unassigned') + '<br>Signals: ' + esc(sig) + '<extra></extra>';
      }),
      showlegend: false
    });
  }
  if (milestoneRows.length) {
    traces.push({
      type: 'scatter',
      mode: 'markers+text',
      x: milestoneRows.map(r => r.task.finish || r.task.start),
      y: milestoneRows.map(r => yById.get(r.task.id)),
      text: milestoneRows.map(r => fmtShort(r.task.finish || r.task.start)),
      textposition: 'middle right',
      marker: {
        size: 12,
        symbol: 'diamond',
        color: milestoneRows.map(r => r.relation === 'ROOT' ? '#4b6bfb' : r.relation === 'PRED' ? '#16a34a' : '#2563eb'),
        line: { width: 1, color: '#1e293b' }
      },
      hovertemplate: milestoneRows.map(r => '<b>' + esc(r.task.name) + '</b><br>Relation: ' + r.relation + ' · Level ' + (r.depth || 0) + '<br>Type: Milestone<br>Date: ' + fmt(r.task.finish || r.task.start) + '<extra></extra>'),
      showlegend: false
    });
  }

  const depSegments = rows
    .filter(r => r.parentTaskId && yById.has(r.parentTaskId))
    .map(r => {
      const parentTask = taskById.get(r.parentTaskId);
      if (!parentTask) return null;
      const fromX = parentTask.finish || parentTask.start;
      const toX = r.task.start || r.task.finish;
      if (!fromX || !toX) return null;
      return {
        fromX,
        fromY: yById.get(parentTask.id),
        toX,
        toY: yById.get(r.task.id)
      };
    })
    .filter(Boolean);

  if (depSegments.length) {
    const x = [];
    const y = [];
    depSegments.forEach(seg => {
      x.push(seg.fromX, seg.toX, null);
      y.push(seg.fromY, seg.toY, null);
    });
    traces.push({
      type: 'scatter',
      mode: 'lines',
      x,
      y,
      line: { color: '#94a3b8', width: 1.4, dash: 'dot' },
      hoverinfo: 'skip',
      showlegend: false
    });
    traces.push({
      type: 'scatter',
      mode: 'markers',
      x: depSegments.map(seg => seg.toX),
      y: depSegments.map(seg => seg.toY),
      marker: {
        size: 9,
        symbol: 'triangle-down-open',
        color: '#ffffff',
        line: { width: 1.4, color: '#94a3b8' }
      },
      hovertemplate: depSegments.map(() => 'Dependency flow ↓<extra></extra>'),
      showlegend: false
    });
  }

  Plotly.react(host, traces, {
    margin: { l: 260, r: 16, t: 10, b: 36 },
    paper_bgcolor: '#ffffff', plot_bgcolor: '#f8fafc',
    height: Math.max(500, rows.length * 32),
    bargap: 0.3,
    uniformtext: { mode: 'hide', minsize: 9 },
    xaxis: { type: 'date', range: [minD.toISOString(), maxD.toISOString()], showgrid: true, gridcolor: '#e2e8f0', gridwidth: 1, tickfont: { color: '#64748b', size: 10 }, zeroline: false },
    yaxis: { automargin: true, autorange: 'reversed', tickfont: { color: '#334155', size: 10.5 } },
    showlegend: false
  }, { displayModeBar: false, responsive: true });
}

function toggleBoardNode(id) {
  if (boardState.expanded.has(id)) boardState.expanded.delete(id);
  else boardState.expanded.add(id);
  renderBoardTree();
}

function toggleBoardMode(mode) {
  // toggle critical or driving; they are mutually exclusive
  if (mode === 'critical') {
    boardState.criticalOnly = !boardState.criticalOnly;
    boardState.drivingOnly = false;
  } else if (mode === 'driving') {
    boardState.drivingOnly = !boardState.drivingOnly;
    boardState.criticalOnly = false;
  }
  const bc = document.getElementById('btnCritical');
  const bd = document.getElementById('btnDriving');
  if (bc) bc.classList.toggle('active', !!boardState.criticalOnly);
  if (bd) bd.classList.toggle('active', !!boardState.drivingOnly);
  if (!boardState.selectedActionId) return;
  selectBoardAction(boardState.selectedActionId);
}
function toggleBoardCriticalOnly(checked) { /* legacy no-op */ }

function selectBoardAction(actionId) {
  boardState.selectedActionId = actionId;
  const action = (DATA.actionItems || []).find(a => a.id === actionId);
  const rootTask = action?.taskId ? taskById.get(action.taskId) : null;
  const list = document.getElementById('boardList');
  if (list) {
    list.querySelectorAll('.board-item').forEach(el => el.classList.remove('active'));
    const selectedEl = document.getElementById('board-item-' + actionId);
    if (selectedEl) selectedEl.classList.add('active');
  }

  if (!rootTask) {
    boardState.rootTask = null;
    boardState.selectedNodeId = null;
    boardState.tree = null;
    boardState.nodeMap = new Map();
    renderBoardTree();
    renderBoardGantt();
    return;
  }

  boardState.rootTask = rootTask;
  boardState.selectedNodeId = '__root__';
  boardState.tree = buildBoardTree(rootTask);
  boardState.nodeMap = boardState.tree.nodeMap;
  boardState.expanded = new Set();
  for (const n of boardState.nodeMap.values()) {
    if (n.depth <= 1) boardState.expanded.add(n.id);
  }
  // summary no longer needed — root card replaces it
  renderBoardTree();
  renderBoardDepDetail(rootTask);
  renderBoardGantt();
}

function renderBoardTab() {
  const list = document.getElementById('boardList');
  const items = (DATA.actionItems || []).slice(0, 120);
  if (!items.length) {
    list.innerHTML = '<div style="color:var(--text2)">No open issues/risks.</div>';
    document.getElementById('boardTree').innerHTML = '<div style="color:var(--text2)">No data.</div>';
    document.getElementById('boardGantt').innerHTML = '<div style="color:var(--text2);padding:10px">No data.</div>';
    renderBoardDepDetail(null);
    return;
  }
  list.innerHTML = items.map(a => {
    const isSlipped = a.taskId ? (taskById.get(a.taskId) || {}).isSlipped : false;
    const chipCls = isSlipped ? ' red' : '';
    return '<div class="board-item" id="board-item-' + a.id + '" onclick="selectBoardAction(\\\'' + a.id + '\\\')">' +
      '<div class="board-item-name">' + (isSlipped ? '🔴 ' : '') + esc(trunc(a.taskName || a.plainIssue, 52)) + '</div>' +
      '<div class="board-item-sub">' +
        (a.workstream ? '<span class="board-item-chip">' + esc(a.workstream) + '</span>' : '') +
        (a.milestoneDate ? '<span class="board-item-chip' + chipCls + '">🏁 ' + esc(fmt(a.milestoneDate)) + '</span>' : '') +
      '</div>' +
      '<div style="font-size:11px;color:var(--text2);margin-top:5px;line-height:1.4">' + esc(trunc(a.plainIssue, 110)) + '</div>' +
    '</div>';
  }).join('');

  const exp = document.getElementById('boardExpandAll');
  const col = document.getElementById('boardCollapseAll');
  if (exp) exp.onclick = () => {
    if (!boardState.nodeMap) return;
    for (const n of boardState.nodeMap.values()) if (n.children.length) boardState.expanded.add(n.id);
    renderBoardTree();
  };
  if (col) col.onclick = () => { boardState.expanded.clear(); renderBoardTree(); };
  // sync button active states
  const bc = document.getElementById('btnCritical');
  const bd = document.getElementById('btnDriving');
  if (bc) bc.classList.toggle('active', !!boardState.criticalOnly);
  if (bd) bd.classList.toggle('active', !!boardState.drivingOnly);

  const first = items.find(a => a.taskId) || items[0];
  if (first) selectBoardAction(first.id);
}

function renderActions() {
  const board = document.getElementById("actionsBoard");
  const sub = document.getElementById("actionsSub");
  const items = DATA.actionItems || [];
  if (!items.length) {
    if (sub) sub.textContent = 'No open issues/risks to summarize';
    board.innerHTML = '<div class="empty">No open risks/issues to summarize.</div>';
    return;
  }

  const nowMs = Date.now();
  const scored = items.slice(0, 300).map(a => {
    const t = a.taskId ? taskById.get(a.taskId) : null;
    const milestoneMs = a.milestoneDate ? new Date(a.milestoneDate).getTime() : null;
    const daysToMilestone = milestoneMs ? Math.round((milestoneMs - nowMs) / 86400000) : null;
    let urgency = 0;
    if (t?.isSlipped) urgency += 5;
    if (t?.isOverdueStart) urgency += 3;
    if (t?.isDue14) urgency += 2;
    if ((a.predCount || 0) + (a.sucCount || 0) >= 4) urgency += 1;
    if (daysToMilestone != null && daysToMilestone <= 14) urgency += 2;
    const lane = urgency >= 6 ? 'immediate' : urgency >= 3 ? 'planned' : 'monitor';
    return { a, t, urgency, lane, daysToMilestone };
  }).sort((x, y) => y.urgency - x.urgency);

  const lanes = [
    { key: 'immediate', label: 'Immediate Resolution', css: 'actions-col-immediate' },
    { key: 'planned', label: 'Planned Resolution', css: 'actions-col-planned' },
    { key: 'monitor', label: 'Monitor Queue', css: 'actions-col-monitor' }
  ];

  if (sub) {
    const totalOpen = scored.length;
    const immediate = scored.filter(x => x.lane === 'immediate').length;
    sub.textContent = totalOpen + ' action cards · ' + immediate + ' immediate';
  }

  const issueTypeIcon = (type, category) => {
    const seed = String(type || '') + ' ' + String(category || '');
    const t = seed.toLowerCase();
    if (/risk/.test(t)) return '⚠️';
    if (/issue/.test(t)) return '🛠️';
    if (/data/.test(t)) return '🗃️';
    if (/test|uat|defect/.test(t)) return '🧪';
    if (/security|access/.test(t)) return '🔐';
    if (/schedule|milestone/.test(t)) return '📅';
    return '📌';
  };

  board.innerHTML = lanes.map(l => {
    const list = scored.filter(x => x.lane === l.key);
    const cards = list.length ? list.map(({ a, t, daysToMilestone }) => {
      const steps = (a.actions || []).slice(0, 3).map(step => '<li>' + esc(step) + '</li>').join('');
      const typeIcon = issueTypeIcon(a.type, a.sourceCategory);
      const typeLabel = (a.type || 'Issue') + (a.sourceCategory ? (' · ' + a.sourceCategory) : '');
      const impactMilestone = a.milestoneName
        ? (a.milestoneName + ' (' + fmt(a.milestoneDate) + ')')
        : (t?.milestone ? (t.name + ' (' + fmt(t.finish || t.start) + ')') : 'No milestone identified');
      const topMeta =
        '<div class="action-top">' +
          '<span class="action-pill ws" title="Workstream for the task with the issue">🏷️ ' + esc(a.workstream || '(Unassigned)') + '</span>' +
          '<span class="action-pill type" title="Issue / risk type">' + typeIcon + ' ' + esc(trunc(typeLabel, 44)) + '</span>' +
          '<span class="action-pill ms" title="Immediate higher-level milestone impacted">🏁 ' + esc(trunc(impactMilestone, 52)) + '</span>' +
        '</div>';
      const dueChip = (daysToMilestone != null)
        ? '<span class="action-chip date">🏁 ' + esc(fmt(a.milestoneDate)) + (daysToMilestone >= 0 ? ' (D-' + daysToMilestone + ')' : ' (late)') + '</span>'
        : '';
      const ownerChip = '<span class="action-chip owner">👤 ' + esc(a.owner || t?.assignedTo || 'Unassigned') + '</span>';
      const depChip = '<span class="action-chip dep">↑' + esc(String(a.predCount || 0)) + ' ↓' + esc(String(a.sucCount || 0)) + '</span>';
      return '<article class="action-card">' +
        topMeta +
        '<div class="action-title">' + esc(trunc(a.plainIssue || a.taskName || 'Issue', 110)) + '</div>' +
        '<div class="action-linked">Task with issue: <strong>' + esc(trunc(a.taskName || 'No linked task', 64)) + '</strong> → Immediate higher-level milestone impacted: <strong>' + esc(trunc(impactMilestone, 64)) + '</strong></div>' +
        '<div class="action-desc">' + esc(trunc(a.plainImpact || a.linkageImpact || 'No impact description available.', 180)) + '</div>' +
        '<div class="action-chips">' + ownerChip + dueChip + depChip + '</div>' +
        (steps ? '<ul class="action-steps">' + steps + '</ul>' : '') +
        '<div class="action-impact">Resolution focus: ' + esc(trunc(a.linkageImpact || 'Track linkage impact and remove blockers.', 120)) + '</div>' +
      '</article>';
    }).join('') : '<div class="empty" style="padding:18px 10px">No items in this lane.</div>';
    return '<section class="actions-col ' + l.css + '">' +
      '<div class="actions-col-hd"><span>' + esc(l.label) + '</span><span>' + list.length + '</span></div>' +
      '<div class="actions-col-body">' + cards + '</div>' +
    '</section>';
  }).join('');
}

function yield_() { return new Promise(r => setTimeout(r, 0)); }
function stageProgress(step, total) {
  const pct = Math.round((step / total) * 100);
  const bar = document.getElementById('loadingBar');
  const pctEl = document.getElementById('loadingPct');
  if (bar) bar.style.width = pct + '%';
  if (pctEl) pctEl.textContent = pct + '%';
  for (let i = 0; i < step; i++) {
    const el = document.getElementById('ls' + i);
    if (el) { el.classList.add('done'); el.classList.remove('active'); el.querySelector('.stage-icon').textContent = '✅'; }
  }
  if (step < total) {
    const active = document.getElementById('ls' + step);
    if (active) { active.classList.add('active'); active.querySelector('.stage-icon').innerHTML = '<span class="loading-spin-sm"></span>'; }
  }
}

// ── Init ───────────────────────────────────────────────────────────────────
async function init() {
  const TOTAL = 5;

  // Stage 0: index
  stageProgress(0, TOTAL);
  await yield_();
  // Update stage 0 label with actual count
  const ls0 = document.getElementById('ls0');
  if (ls0) ls0.innerHTML = '<span class="stage-icon"><span class="loading-spin-sm"></span></span> Indexing ' + DATA.tasks.length + ' tasks…';
  await yield_();

  // Stage 1: KPIs and status chart
  stageProgress(1, TOTAL);
  await yield_();
  const m = DATA.metrics;
  document.getElementById("healthBadge").textContent = m.healthLabel;
  document.getElementById("genTime").textContent = "Generated " + new Date(DATA.generatedAt).toLocaleString();
  const kpis = [
    {label:"Total Tasks",val:m.total,cls:"blue"},
    {label:"Open",val:m.totalOpen,cls:"blue"},
    {label:"Slipped Open",val:m.slippedOpen,cls:"red"},
    {label:"Slip Rate",val:m.slipRatePct+"%",cls:parseFloat(m.slipRatePct)>=75?"red":parseFloat(m.slipRatePct)>=50?"amber":"green"},
    {label:"Due 14d",val:m.due14,cls:m.due14>=100?"red":m.due14>=60?"amber":"green"},
    {label:"Overdue Starts",val:m.overdueStarts,cls:m.overdueStarts>=100?"red":m.overdueStarts>=40?"amber":"green"},
    {label:"Open Issues",val:m.openIssues,cls:m.openIssues>5?"amber":"green"},
    {label:"Open Risks",val:m.openRisks,cls:m.openRisks>5?"amber":"green"},
  ];
  const strip = document.getElementById("kpiStrip");
  kpis.forEach(k => {
    const el = document.createElement("div");
    el.className = "kpi "+k.cls;
    el.innerHTML = \`<span class="val">\${esc(String(k.val))}</span><span class="lbl">\${esc(k.label)}</span>\`;
    strip.appendChild(el);
  });
  const statusColors = ['#7b8cff','#5aa9e6','#d85b72','#c7922a','#2f8f6b','#99a8b8'];
  new Chart(document.getElementById("statusChart"),{type:"doughnut",data:{labels:DATA.statusBreakdown.map(x=>x.s),datasets:[{data:DATA.statusBreakdown.map(x=>x.c),backgroundColor:DATA.statusBreakdown.map((_,i)=>statusColors[i%statusColors.length]),borderColor:'#ffffff',borderWidth:2}]},options:{responsive:true,maintainAspectRatio:false,plugins:{legend:{position:"bottom",labels:{color:"#5f7185",font:{size:10},boxWidth:10,padding:6}}}}});

  // Stage 2: milestones
  stageProgress(2, TOTAL);
  await yield_();
  renderMilestoneMini();

  // Stage 3: risk/issue tiles
  stageProgress(3, TOTAL);
  await yield_();
  const typeList = document.getElementById("typeList");
  DATA.riskTypeBreakdown.forEach(rt => {
    const tile = document.createElement("div");
    tile.className = "type-tile"; tile.dataset.typeKey = rt.typeKey;
    tile.innerHTML = \`<span class="type-label">\${esc(rt.typeKey)}</span><span class="type-badge">\${rt.count}</span>\`;
    tile.addEventListener("click", () => toggleFilter(rt.typeKey, tile));
    typeList.appendChild(tile);
  });
  const exp = document.getElementById('planExpandAll');
  const col = document.getElementById('planCollapseAll');
  if (exp && col) {
    exp.addEventListener('click', () => {
      if (!planState) planState = buildPlanState();
      planState.byId.forEach(n => { if (n.children.length) planState.expanded.add(n.id); });
      renderPlanExplorer();
    });
    col.addEventListener('click', () => {
      if (!planState) planState = buildPlanState();
      planState.expanded.clear();
      renderPlanExplorer();
    });
  }

  // Stage 4: task table
  stageProgress(4, TOTAL);
  await yield_();
  const tbody = document.getElementById("taskBody");
  if (tbody) {
    tbody.innerHTML = '<tr><td colspan="11" class="empty">Loading at-risk tasks…</td></tr>';
  }
  stageProgress(TOTAL, TOTAL);
  await yield_();
  hideLoading();
  scheduleNonBlocking(() => renderTasks(PERF.initialTaskRows));
}

// ── Tasks tab ─────────────────────────────────────────────────────────────
function toggleFilter(typeKey, tile) {
  if (activeFilter === typeKey) { clearFilter(); return; }
  activeFilter = typeKey;
  document.querySelectorAll(".type-tile").forEach(t => t.classList.remove("active"));
  tile.classList.add("active");
  document.getElementById("filterClear").style.display = "inline";
  renderTasks(PERF.maxTaskRows);
}
function clearFilter() {
  activeFilter = null;
  document.querySelectorAll(".type-tile").forEach(t => t.classList.remove("active"));
  document.getElementById("filterClear").style.display = "none";
  renderTasks(PERF.maxTaskRows);
}
function renderTasks(maxRows = PERF.maxTaskRows) {
  let tasks = DATA.tasks;
  if (activeFilter) {
    tasks = tasks.filter(t => t.linkedRisks.some(r => r.type+" — "+r.category === activeFilter));
    const shown = Math.min(tasks.length, maxRows);
    document.getElementById("filterInfo").textContent = \`\${tasks.length} tasks linked to "\${activeFilter}" · showing \${shown}\`;
  } else {
    if (!atRiskTasksCache) {
      atRiskTasksCache = DATA.tasks
        .filter(t => t.isSlipped || t.isOverdueStart || t.isDue14 || t.linkedRisks.length > 0)
        .sort((a,b) => (b.slipDays||0)-(a.slipDays||0));
    }
    tasks = atRiskTasksCache;
    const shown = Math.min(tasks.length, maxRows);
    document.getElementById("filterInfo").textContent = \`\${tasks.length} at-risk tasks · showing \${shown}\`;
  }
  const tbody = document.getElementById("taskBody");
  if (!tasks.length) { tbody.innerHTML = '<tr><td colspan="11" class="empty">No matching tasks.</td></tr>'; return; }
  tbody.innerHTML = tasks.slice(0, maxRows).map(t => {
    const pills = t.linkedRisks.map(r => \`<span class="pill \${/issue/i.test(r.type)?"issue":"risk"}" title="\${esc(r.severity)}">\${esc(trunc(r.name,26))}</span>\`).join("");
    const slipCell = t.slipDays != null ? \`<span class="pill slip">+\${t.slipDays}d</span>\` : t.pct >= 100 ? \`<span class="pill done">Done</span>\` : "—";
    return \`<tr class="\${selectedTaskId===t.id?" selected":""}" onclick="selectTask(event,'\${t.id}')">
      <td style="color:var(--text2);white-space:nowrap">\${t.uid||"—"}</td>
      <td><div class="t-name">\${esc(trunc(t.name,55))}</div><div class="t-ws">\${esc(t.workstream)}</div></td>
      <td style="white-space:nowrap">\${t.milestone ? '<span class="pill done">🏁 Yes</span>' : '—'}</td>
      <td style="white-space:nowrap">\${esc(t.status)}</td>
      <td style="white-space:nowrap">\${fmt(t.start)}</td>
      <td style="white-space:nowrap">\${fmt(t.finish)}</td>
      <td style="white-space:nowrap">\${fmt(t.baselineFinish)}</td>
      <td>\${slipCell}</td>
      <td style="color:var(--text2)">\${t.predCount||"—"}</td>
      <td style="color:var(--text2)">\${t.sucCount||"—"}</td>
      <td>\${pills||'<span style="color:var(--text2)">—</span>'}</td>
    </tr>\`;
  }).join("");
}
function selectTask(evt, taskId) {
  if (selectedTaskId === taskId && document.getElementById("depPanel").classList.contains("open")) {
    closeDepPanel();
    return;
  }
  selectedTaskId = taskId;
  document.querySelectorAll("#taskBody tr").forEach(r => r.classList.remove("selected"));
  evt.currentTarget.classList.add("selected");
  openDepPanel(taskId);
}
function openDepPanel(taskId) {
  const task = taskById.get(taskId); if (!task) return;
  document.getElementById("depTitle").textContent = trunc(task.name, 38);
  document.getElementById("depMeta").innerHTML = task.milestone ? '<span class="dep-head-badge">🏁 Milestone</span>' : '';
  document.getElementById("depPanel").classList.add("open");
  renderDepPanel(task);
}
function closeDepPanel() {
  document.getElementById("depPanel").classList.remove("open");
  document.getElementById("depMeta").innerHTML = "";
  selectedTaskId = null;
  document.querySelectorAll("#taskBody tr").forEach(r => r.classList.remove("selected"));
}
function depBox(t, role) {
  if (!t) return "";
  const slipped = t.slipDays != null && t.slipDays > 0;
  return \`<div class="dep-box \${role}\${slipped?" slipped":""}">
    <div class="box-name">\${esc(trunc(t.name,28))}</div>
    <div class="box-meta">#\${t.uid||"?"} · \${esc(t.workstream)}</div>
    <div class="box-meta">Finish: \${fmt(t.finish)}</div>
    \${slipped?\`<div class="box-slip">+\${t.slipDays}d slip</div>\`:""}
    \${t.pct>=100?\`<div class="box-slip" style="color:var(--green)">✓ Done</div>\`:""}
  </div>\`;
}
function renderDepPanel(task) {
  const preds = (task.predIds||[]).map(id => taskById.get(id)).filter(Boolean);
  const sucs  = (task.sucIds||[]).map(id => taskById.get(id)).filter(Boolean);
  const slipped = task.slipDays != null && task.slipDays > 0;
  const predSlipped = preds.filter(p => p.slipDays != null && p.slipDays > 0);
  const sucSlipped  = sucs.filter(s => s.slipDays != null && s.slipDays > 0);
  let summaryColor = "var(--green)", summaryText = "No upstream slip detected.";
  if (predSlipped.length > 0 && slipped) { summaryColor="var(--red)"; summaryText=\`\${predSlipped.length} upstream predecessor(s) are late — this task is also slipped.\`; }
  else if (predSlipped.length > 0) { summaryColor="var(--amber)"; summaryText=\`\${predSlipped.length} predecessor(s) are late — this task may slip next.\`; }
  else if (slipped) { summaryColor="var(--amber)"; summaryText=\`This task is slipped (+\${task.slipDays}d). \${sucs.length} downstream successor(s) at risk.\`; }
  const predCol = preds.length ? preds.map(p=>depBox(p,"pred")).join("") : \`<div class="dep-box pred" style="opacity:.4;font-style:italic">No predecessors</div>\`;
  const sucCol  = sucs.length  ? sucs.map(s=>depBox(s,"suc")).join("")  : \`<div class="dep-box suc"  style="opacity:.4;font-style:italic">No successors</div>\`;
  const issueRows = task.linkedRisks.map(r => \`<div class="dep-detail-row"><div class="dr-label">\${esc(r.type)} — \${esc(r.category)}</div><div>\${esc(r.name)}</div><div style="color:var(--amber);font-size:10px;margin-top:2px">Severity: \${esc(r.severity)} · Status: \${esc(r.status||"Open")}</div></div>\`).join("");
  document.getElementById("depBody").innerHTML = \`
    <div style="background:var(--surface2);border-radius:6px;padding:8px 10px;margin-bottom:12px;font-size:11px;border-left:3px solid \${summaryColor}">\${summaryText}</div>
    <div class="dep-section"><h5>Dependency Flow</h5>
      <div class="dep-flow">
        <div class="dep-col"><div style="font-size:9px;text-transform:uppercase;color:var(--text2);letter-spacing:.5px;margin-bottom:4px">↑ Predecessors (\${preds.length})</div>\${predCol}</div>
        <div class="dep-arrows">\${preds.length?'<div class="arr">→</div>':""}</div>
        <div class="dep-col center"><div style="font-size:9px;text-transform:uppercase;color:var(--text2);letter-spacing:.5px;margin-bottom:4px">● This Task</div>\${depBox(task,"focal")}</div>
        <div class="dep-arrows">\${sucs.length?'<div class="arr">→</div>':""}</div>
        <div class="dep-col"><div style="font-size:9px;text-transform:uppercase;color:var(--text2);letter-spacing:.5px;margin-bottom:4px">↓ Successors (\${sucs.length})</div>\${sucCol}</div>
      </div>
    </div>
    \${predSlipped.length>0?\`<div class="dep-section"><h5>⚠ Slipped Predecessors (\${predSlipped.length})</h5>\${predSlipped.map(p=>\`<div class="dep-detail-row"><div class="dr-label">#\${p.uid||"?"} · \${esc(p.workstream)}</div><div>\${esc(trunc(p.name,50))}</div><div style="color:var(--amber);font-size:10px;margin-top:2px">Slip +\${p.slipDays}d · Finish: \${fmt(p.finish)}</div></div>\`).join("")}</div>\`:""}
    \${sucSlipped.length>0?\`<div class="dep-section"><h5>⚠ Slipped Successors (\${sucSlipped.length})</h5>\${sucSlipped.map(s=>\`<div class="dep-detail-row"><div class="dr-label">#\${s.uid||"?"} · \${esc(s.workstream)}</div><div>\${esc(trunc(s.name,50))}</div><div style="color:var(--amber);font-size:10px;margin-top:2px">Slip +\${s.slipDays}d · Finish: \${fmt(s.finish)}</div></div>\`).join("")}</div>\`:""}
    \${task.linkedRisks.length>0?\`<div class="dep-section"><h5>Linked Issues / Risks (\${task.linkedRisks.length})</h5>\${issueRows}</div>\`:""}
    <div class="dep-section"><h5>Task Details</h5>
      <div class="dep-detail-row"><div class="dr-label">Workstream</div><div>\${esc(task.workstream)}</div></div>
      <div class="dep-detail-row"><div class="dr-label">Start → Finish</div><div>\${fmt(task.start)} → \${fmt(task.finish)}</div>\${task.baselineFinish?\`<div style="color:var(--text2);font-size:10px;margin-top:2px">Baseline: \${fmt(task.baselineFinish)}</div>\`:""}</div>
      <div class="dep-detail-row"><div class="dr-label">% Complete</div><div>\${task.pct}%</div></div>
      \${task.milestone?\`<div class="dep-detail-row" style="color:var(--amber)">🏁 Milestone</div>\`:""}
    </div>\`;
}

// ── Gantt tab ──────────────────────────────────────────────────────────────
function renderGantt() {
  const wsFilter = document.getElementById("ganttWs").value;
  const timeWindow = document.getElementById("ganttWindow").value;
  const scale = document.getElementById("ganttScale").value;
  const riskFilter = document.getElementById("ganttRiskType").value;
  const showMiles = document.getElementById("ganttMilestones").checked;
  const showIssue = document.getElementById("ganttIssues").checked;
  const showSlip = document.getElementById("ganttSlipped").checked;

  const hasRiskType = (t, key) => t.linkedRisks.some(r => (r.type + " — " + r.category) === key);

  let selected = DATA.tasks.filter(t => {
    if (t.pct >= 100) return false;
    if (wsFilter && t.workstream !== wsFilter) return false;

    if (riskFilter === "__any__" && t.linkedRisks.length === 0) return false;
    if (riskFilter && riskFilter !== "__any__" && !hasRiskType(t, riskFilter)) return false;

    if (!riskFilter) {
      const match = (showMiles && t.milestone) || (showIssue && t.linkedRisks.length > 0) || (showSlip && t.isSlipped);
      if (!match) return false;
    }
    return true;
  });

  const milestones = showMiles ? selected.filter(t => t.milestone) : [];
  const tasks = selected.filter(t => !t.milestone).sort((a, b) => (a.finish || "").localeCompare(b.finish || ""));
  const visibleIds = new Set(tasks.map(t => t.id));

  const countLabel = tasks.length + " tasks" + (milestones.length ? (" + " + milestones.length + " milestones") : "");
  document.getElementById("ganttCount").textContent = countLabel;
  document.getElementById("ganttScaleNote").textContent = "Grouped by workstream → risk type • scale: " + scale;

  if (!tasks.length && !milestones.length) {
    document.getElementById("ganttPlot").innerHTML = '<div style="padding:20px;color:var(--text2);font-size:12px">No tasks match filters.</div>';
    return;
  }

  const datePool = [].concat(tasks, milestones)
    .flatMap(t => [t.start, t.finish, t.baselineFinish].filter(Boolean).map(d => new Date(d).getTime()))
    .filter(n => !isNaN(n));

  let minDate = new Date(Math.min.apply(null, datePool));
  let maxDate = new Date(Math.max.apply(null, datePool));
  minDate.setDate(1);
  maxDate.setMonth(maxDate.getMonth() + 1, 1);

  if (timeWindow !== "auto") {
    const now = new Date();
    now.setDate(1);
    minDate = new Date(now);
    const monthsOut = timeWindow === "3m" ? 3 : timeWindow === "6m" ? 6 : 12;
    maxDate = new Date(now);
    maxDate.setMonth(maxDate.getMonth() + monthsOut, 1);
  }

  const totalDays = Math.max(1, (maxDate - minDate) / 86400000);
  const groupedTasks = [];
  const workstreams = [...new Set(tasks.map(t => t.workstream))].sort();
  workstreams.forEach(ws => {
    groupedTasks.push({ kind: "lane", label: ws, workstream: ws });
    const inWs = tasks.filter(t => t.workstream === ws);
    const riskGroups = [...new Set(inWs.map(topRiskKey))].sort();
    riskGroups.forEach(group => {
      groupedTasks.push({ kind: "riskHeader", label: group, workstream: ws });
      inWs.filter(t => topRiskKey(t) === group).forEach(t => groupedTasks.push({ kind: "task", task: t }));
    });
  });
  const sevScore = s => {
    const v = String(s || "").toLowerCase();
    if (/(critical|very high|urgent|severe)/.test(v)) return 4;
    if (/high/.test(v)) return 3;
    if (/medium|med/.test(v)) return 2;
    if (/low/.test(v)) return 1;
    return 0;
  };

  const riskIcon = t => {
    if (!t.linkedRisks.length) return "";
    let top = t.linkedRisks[0];
    t.linkedRisks.forEach(r => { if (sevScore(r.severity) > sevScore(top.severity)) top = r; });
    const type = String(top.type || "").toLowerCase();
    const core = /issue/.test(type) ? "🛠" : /risk/.test(type) ? "⚠" : "•";
    const sev = sevScore(top.severity);
    const sevMark = sev >= 3 ? "🔴" : sev === 2 ? "🟠" : "🟡";
    return core + sevMark;
  };
  const rows = [];
  if (showMiles && milestones.length) rows.push({ kind: 'milestoneHeader', label: '🏁 Milestones' });
  milestones.forEach(m => rows.push({ kind: 'milestone', label: '  ◆ ' + trunc(m.name, 28), task: m }));
  groupedTasks.forEach(entry => {
    if (entry.kind === 'lane') rows.push({ kind: 'lane', label: '🧭 ' + entry.label, workstream: entry.workstream });
    else if (entry.kind === 'riskHeader') rows.push({ kind: 'riskHeader', label: '   ⚑ ' + entry.label, workstream: entry.workstream });
    else rows.push({ kind: 'task', label: '      • ' + trunc(entry.task.name, 36), task: entry.task });
  });

  const yOrder = rows.map(r => r.label);
  const taskRows = rows.filter(r => r.kind === 'task');
  const milestoneRows = rows.filter(r => r.kind === 'milestone');
  const laneRows = rows.filter(r => r.kind === 'lane');
  const riskRows = rows.filter(r => r.kind === 'riskHeader');
  const milestoneHeaderRows = rows.filter(r => r.kind === 'milestoneHeader');

  const rangeMs = Math.max(86400000, maxDate.getTime() - minDate.getTime());
  const fullSpan = new Array(laneRows.length).fill(rangeMs);
  const riskSpan = new Array(riskRows.length).fill(rangeMs);
  const milestoneHeaderSpan = new Array(milestoneHeaderRows.length).fill(rangeMs);
  const taskDurations = taskRows.map(r => Math.max(86400000, new Date(r.task.finish).getTime() - new Date(r.task.start || r.task.finish).getTime() + 86400000));

  const taskText = taskRows.map(r => {
    const icon = riskIcon(r.task);
    const dep = ((r.task.predIds || []).length || (r.task.sucIds || []).length) ? ' ↑' + (r.task.predIds || []).length + ' ↓' + (r.task.sucIds || []).length : '';
    return fmtShort(r.task.start || r.task.finish) + ' → ' + fmtShort(r.task.finish) + (icon ? '  ' + icon : '') + dep;
  });

  const traces = [];
  if (milestoneHeaderRows.length) {
    traces.push({
      type: 'bar',
      orientation: 'h',
      y: milestoneHeaderRows.map(r => r.label),
      base: milestoneHeaderRows.map(() => minDate),
      x: milestoneHeaderSpan,
      marker: { color: '#ecfdf5', line: { color: '#b7e5cf', width: 1 } },
      hoverinfo: 'skip',
      showlegend: false,
      textposition: 'none'
    });
  }
  if (laneRows.length) {
    traces.push({
      type: 'bar',
      orientation: 'h',
      y: laneRows.map(r => r.label),
      base: laneRows.map(() => minDate),
      x: fullSpan,
      marker: { color: '#eaf1ff', line: { color: '#c6d7ff', width: 1 } },
      hoverinfo: 'skip',
      showlegend: false,
      textposition: 'none'
    });
  }
  if (riskRows.length) {
    traces.push({
      type: 'bar',
      orientation: 'h',
      y: riskRows.map(r => r.label),
      base: riskRows.map(() => minDate),
      x: riskSpan,
      marker: { color: '#f8fafc', line: { color: '#d7e0ea', width: 1 } },
      hoverinfo: 'skip',
      showlegend: false,
      textposition: 'none'
    });
  }
  if (taskRows.length) {
    traces.push({
      type: 'bar',
      orientation: 'h',
      y: taskRows.map(r => r.label),
      base: taskRows.map(r => r.task.start || r.task.finish),
      x: taskDurations,
      marker: {
        color: taskRows.map(r => r.task.isSlipped ? '#d97757' : r.task.pct > 0 ? '#5b7cff' : '#8fa1b5'),
        line: { color: taskRows.map(r => r.task.isSlipped ? '#b45b42' : r.task.pct > 0 ? '#3f5de0' : '#6f8196'), width: 1 }
      },
      text: taskText,
      textposition: 'inside',
      insidetextanchor: 'middle',
      textfont: { color: '#ffffff', size: 10 },
      hovertemplate: taskRows.map(r => '<b>' + esc(r.task.name) + '</b><br>Workstream: ' + esc(r.task.workstream) + '<br>Start: ' + fmt(r.task.start) + '<br>Finish: ' + fmt(r.task.finish) + (r.task.baselineFinish ? '<br>Baseline: ' + fmt(r.task.baselineFinish) : '') + (r.task.linkedRisks.length ? '<br>Linked: ' + r.task.linkedRisks.length + ' risk(s)/issue(s)' : '') + '<extra></extra>'),
      showlegend: false,
      cliponaxis: false
    });
  }
  if (milestoneRows.length) {
    traces.push({
      type: 'scatter',
      mode: 'markers+text',
      x: milestoneRows.map(r => r.task.finish || r.task.start),
      y: milestoneRows.map(r => r.label),
      text: milestoneRows.map(r => (r.task.linkedRisks.length ? '⚠ ' : '') + fmtShort(r.task.finish || r.task.start)),
      textposition: 'middle right',
      marker: {
        size: 13,
        symbol: 'diamond',
        color: milestoneRows.map(r => r.task.isSlipped ? '#d85b72' : '#2f8f6b'),
        line: { width: 1, color: milestoneRows.map(r => r.task.isSlipped ? '#b43b55' : '#256f54') }
      },
      hovertemplate: milestoneRows.map(r => '<b>' + esc(r.task.name) + '</b><br>Milestone<br>Date: ' + fmt(r.task.finish || r.task.start) + '<extra></extra>'),
      showlegend: false
    });
  }

  const layout = {
    margin: { l: 250, r: 26, t: 18, b: 40 },
    paper_bgcolor: '#ffffff',
    plot_bgcolor: '#ffffff',
    barmode: 'overlay',
    bargap: 0.28,
    height: Math.max(680, rows.length * 28 + 60),
    xaxis: {
      type: 'date',
      range: [minDate.toISOString(), maxDate.toISOString()],
      showgrid: true,
      gridcolor: '#e7edf4',
      tickfont: { color: '#5f7185', size: 10 },
      zeroline: false,
      tickformat: scale === 'months' ? '%b %y' : '%b %-d',
      dtick: scale === 'months' ? 'M1' : 7 * 24 * 60 * 60 * 1000
    },
    yaxis: {
      automargin: true,
      categoryorder: 'array',
      categoryarray: yOrder,
      autorange: 'reversed',
      tickfont: { color: '#334155', size: 11 },
      showgrid: false
    },
    shapes: [{
      type: 'line',
      x0: new Date().toISOString().slice(0, 10),
      x1: new Date().toISOString().slice(0, 10),
      y0: 0,
      y1: 1,
      yref: 'paper',
      line: { color: '#d85b72', width: 2, dash: 'dot' }
    }],
    annotations: [{
      x: new Date().toISOString().slice(0, 10),
      y: 1.05,
      yref: 'paper',
      text: 'Today',
      showarrow: false,
      font: { color: '#b43b55', size: 10 }
    }],
    showlegend: false
  };

  Plotly.react(document.getElementById('ganttPlot'), traces, layout, { displayModeBar: false, responsive: true });
}

init().catch(e => console.error('Dashboard init failed:', e));
<\/script>
</body>
</html>`;
}

main().catch(e => { console.error(e); process.exit(1); });

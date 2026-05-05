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
      resourceNames,
      businessOwners,
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
  const actionItems = riskRecords
    .map(r => {
      const linkedTasks = r.linkedTaskIds.map(id => taskById.get(id)).filter(Boolean);
      const aiCategory = copilotCategory(r);
      const plainIssue = simpleSentence(r.name);
      const plainImpact = copilotImpact(r.severity, linkedTasks.length, r.type);
      const actions = copilotActions(r, linkedTasks);
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
  const chartScript = CHARTJS_INLINE
    ? `<script>${CHARTJS_INLINE.replace(/<\/script>/gi, "<\\/script>")}<\/script>`
    : `<script src="https://cdn.jsdelivr.net/npm/chart.js@4.4.3/dist/chart.umd.min.js"><\/script>`;
  const plotlyScript = PLOTLY_INLINE
    ? `<script>${PLOTLY_INLINE.replace(/<\/script>/gi, "<\\/script>")}<\/script>`
    : `<script src="https://cdn.jsdelivr.net/npm/plotly.js-dist-min@2.35.2/plotly.min.js"><\/script>`;

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
  ::-webkit-scrollbar{width:5px;height:5px;} ::-webkit-scrollbar-track{background:transparent;} ::-webkit-scrollbar-thumb{background:var(--border);border-radius:3px;}
</style>
</head>
<body>
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
    <div class="panel-title">Copilot Insights — Simplified issue, impact, and recommended resolution</div>
    <div class="tbl-wrap" style="height:100%;box-shadow:none;border-radius:10px;">
      <table>
        <thead>
          <tr>
            <th>AI Category</th><th>Issue (Plain)</th><th>Impact (Plain)</th><th>Recommended Actions</th><th>Milestones + Assigned</th>
          </tr>
        </thead>
        <tbody id="actionBody"></tbody>
      </table>
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
const taskById = new Map(DATA.tasks.map(t => [t.id, t]));
const taskByUid = new Map(DATA.tasks.filter(t => t.uid != null).map(t => [t.uid, t]));
let planState = null;
const fmt = d => d ? new Date(d).toLocaleDateString("en-US",{month:"short",day:"numeric",year:"2-digit"}) : "—";
const fmtShort = d => d ? new Date(d).toLocaleDateString("en-US",{month:"short",day:"numeric"}) : "—";
const esc = s => String(s||"").replace(/&/g,"&amp;").replace(/</g,"&lt;").replace(/>/g,"&gt;");
const trunc = (s,n) => s && s.length>n ? s.slice(0,n)+"…" : (s||"");
const riskKey = r => (r.type || "Risk") + " — " + (r.category || "(Uncategorized)");
const topRiskKey = task => task.linkedRisks.length ? riskKey(task.linkedRisks[0]) : "No linked risk";

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

  return { byId, roots, expanded };
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
  const n = planState.byId.get(nodeId);
  if (!n) return false;
  const selfIssue = (n.task.linkedRisks || []).some(r => /issue/i.test(r.type));
  if (selfIssue) return true;
  for (const cid of n.children) {
    if (hasIssueRollup(cid)) return true;
  }
  return false;
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
            ? '<button onclick="togglePlanNode(\'' + n.id + '\')" style="border:1px solid var(--border);background:var(--surface2);border-radius:4px;width:18px;height:18px;cursor:pointer;font-size:11px;line-height:1">' + (isOpen ? '−' : '+') + '</button>'
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

  renderPlanGantt(rows);
}

function togglePlanNode(id) {
  if (!planState) return;
  if (planState.expanded.has(id)) planState.expanded.delete(id);
  else planState.expanded.add(id);
  renderPlanExplorer();
}

function renderPlanGantt(rows) {
  const host = document.getElementById('planGantt');
  const taskRows = rows.filter(r => r.node.task.pct < 100 && (r.node.start || r.node.finish)).slice(0, 220);
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

  const trace = {
    type: 'scatter',
    mode: 'markers+text',
    x: milestones.map(m => m.finish || m.start),
    y: milestones.map(m => trunc(m.name, 16)),
    text: milestones.map(m => (m.linkedRisks.length ? '⚠ ' : '◆ ') + trunc(m.assignedTo || 'Unassigned', 10)),
    textposition: 'middle right',
    marker: {
      size: 12,
      symbol: 'diamond',
      color: milestones.map(m => m.linkedRisks.length ? '#d85b72' : '#2f8f6b'),
      line: { width: 1, color: milestones.map(m => m.linkedRisks.length ? '#b43b55' : '#256f54') }
    },
    hovertemplate: milestones.map(m => '<b>' + esc(m.name) + '</b><br>Assigned: ' + esc(m.assignedTo || 'Unassigned') + '<br>Date: ' + fmt(m.finish || m.start) + '<extra></extra>')
  };
  const layout = {
    margin: { l: 82, r: 8, t: 8, b: 18 },
    paper_bgcolor: '#ffffff',
    plot_bgcolor: '#ffffff',
    xaxis: {
      type: 'date',
      showgrid: true,
      gridcolor: '#e7edf4',
      tickfont: { size: 9, color: '#5f7185' },
      range: [from.toISOString(), to.toISOString()],
      zeroline: false
    },
    yaxis: {
      automargin: true,
      tickfont: { size: 9, color: '#334155' },
      autorange: 'reversed'
    },
    showlegend: false,
    height: 140
  };
  Plotly.react(host, [trace], layout, { displayModeBar: false, responsive: true, staticPlot: true });
}

// ── Tab switching ──────────────────────────────────────────────────────────
function switchTab(tab) {
  const isGantt = tab === "gantt";
  const isActions = tab === "actions";
  const isPlan = tab === "plan";
  document.getElementById("tabTasks").classList.toggle("active", !isGantt && !isActions && !isPlan);
  document.getElementById("tabGantt").classList.toggle("active", isGantt);
  document.getElementById("tabActions").classList.toggle("active", isActions);
  document.getElementById("tabPlan").classList.toggle("active", isPlan);
  const tc = document.getElementById("tasksContent");
  tc.style.display = isGantt || isActions || isPlan ? "none" : "contents";
  const ac = document.getElementById("actionsContent");
  ac.style.display = isActions ? "flex" : "none";
  const gc = document.getElementById("ganttContent");
  gc.style.display = isGantt ? "flex" : "none";
  const pc = document.getElementById("planContent");
  pc.style.display = isPlan ? "flex" : "none";
  if (isGantt && !ganttRendered) { renderGantt(); ganttRendered = true; }
  if (isActions) renderActions();
  if (isPlan && !planRendered) {
    renderPlanCalendar();
    renderPlanExplorer();
    planRendered = true;
  }
}

function renderActions() {
  const body = document.getElementById("actionBody");
  const items = DATA.actionItems || [];
  if (!items.length) {
    body.innerHTML = '<tr><td colspan="5" class="empty">No open risks/issues to summarize.</td></tr>';
    return;
  }
  body.innerHTML = items.slice(0, 250).map(a => {
    const acts = (a.actions || []).map(x => '<div>• ' + esc(x) + '</div>').join('');
    const milestones = (a.milestoneOwners || []).length
      ? a.milestoneOwners.map(m => '<div><span class="pill done">🏁</span> ' + esc(trunc(m.name, 28)) + ' — ' + esc(m.assignedTo || 'Unassigned') + ' (' + fmt(m.finish) + ')</div>').join('')
      : '<span style="color:var(--text2)">No linked milestones</span>';
    return '<tr>' +
      '<td><span class="pill issue">' + esc(a.aiCategory) + '</span><div style="font-size:10px;color:var(--text2);margin-top:4px">' + esc(a.type) + ' · ' + esc(a.severity || '(Unrated)') + '</div></td>' +
      '<td>' + esc(a.plainIssue) + '</td>' +
      '<td>' + esc(a.plainImpact) + '</td>' +
      '<td>' + acts + '</td>' +
      '<td>' + milestones + '</td>' +
      '</tr>';
  }).join('');
}

// ── Init ───────────────────────────────────────────────────────────────────
function init() {
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
  renderMilestoneMini();
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
  renderTasks();
}

// ── Tasks tab ─────────────────────────────────────────────────────────────
function toggleFilter(typeKey, tile) {
  if (activeFilter === typeKey) { clearFilter(); return; }
  activeFilter = typeKey;
  document.querySelectorAll(".type-tile").forEach(t => t.classList.remove("active"));
  tile.classList.add("active");
  document.getElementById("filterClear").style.display = "inline";
  renderTasks();
}
function clearFilter() {
  activeFilter = null;
  document.querySelectorAll(".type-tile").forEach(t => t.classList.remove("active"));
  document.getElementById("filterClear").style.display = "none";
  renderTasks();
}
function renderTasks() {
  let tasks = DATA.tasks;
  if (activeFilter) {
    tasks = tasks.filter(t => t.linkedRisks.some(r => r.type+" — "+r.category === activeFilter));
    document.getElementById("filterInfo").textContent = \`\${tasks.length} tasks linked to "\${activeFilter}"\`;
  } else {
    tasks = tasks.filter(t => t.isSlipped || t.isOverdueStart || t.isDue14 || t.linkedRisks.length > 0);
    tasks.sort((a,b) => (b.slipDays||0)-(a.slipDays||0));
    document.getElementById("filterInfo").textContent = \`\${tasks.length} at-risk tasks\`;
  }
  const tbody = document.getElementById("taskBody");
  if (!tasks.length) { tbody.innerHTML = '<tr><td colspan="11" class="empty">No matching tasks.</td></tr>'; return; }
  tbody.innerHTML = tasks.slice(0,400).map(t => {
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

init();
<\/script>
</body>
</html>`;
}

main().catch(e => { console.error(e); process.exit(1); });

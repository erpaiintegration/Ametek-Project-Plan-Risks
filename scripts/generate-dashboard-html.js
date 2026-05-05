/* eslint-disable no-console */
require("dotenv/config");
const { Client } = require("@notionhq/client");
const fs = require("fs");
const path = require("path");

// Inline Chart.js so the HTML works offline / file:// / Notion embed
const CHARTJS_PATH = path.join(__dirname, "..", "node_modules", "chart.js", "dist", "chart.umd.js");
const CHARTJS_INLINE = fs.existsSync(CHARTJS_PATH) ? fs.readFileSync(CHARTJS_PATH, "utf8") : null;
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
function relCountByNames(props, names) {
  for (const n of names) { const p = props[n]; if (p?.type === "relation") return relCount(p); }
  return 0;
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
    const predCount = relCountByNames(p, ["Predecessor Tasks", "Predecessors"]);
    const sucCount = relCountByNames(p, ["Successor Tasks", "Successors"]);
    const predIds = [...relIds(p["Predecessor Tasks"]), ...relIds(p.Predecessors)];
    const sucIds = [...relIds(p["Successor Tasks"]), ...relIds(p.Successors)];

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
      predCount,
      sucCount,
      predIds,
      sucIds,
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

  const payload = {
    generatedAt: now.toISOString(),
    metrics: { total, totalOpen, totalDone, slippedOpen, slipRatePct, overdueStarts, due14: due14Count, openIssues: openIssues.length, openRisks: openRisks.length, healthLabel },
    riskTypeBreakdown,
    statusBreakdown,
    topWorkstreams,
    tasks,
    riskRecords
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
  const chartScript = CHARTJS_INLINE
    ? `<script>${CHARTJS_INLINE.replace(/<\/script>/gi, "<\\/script>")}<\/script>`
    : `<script src="https://cdn.jsdelivr.net/npm/chart.js@4.4.3/dist/chart.umd.min.js"><\/script>`;

  return `<!DOCTYPE html>
<html lang="en">
<head>
<meta charset="UTF-8"/>
<meta name="viewport" content="width=device-width,initial-scale=1"/>
<title>Ametek SAP S4 — Schedule Dashboard</title>
${chartScript}
<style>
  :root{--bg:#0f1117;--surface:#1a1d27;--surface2:#22263a;--border:#2e3350;--text:#e8eaf6;--text2:#9096b8;--red:#ef5350;--amber:#ff9800;--green:#66bb6a;--blue:#42a5f5;--accent:#5c6bc0;}
  *{box-sizing:border-box;margin:0;padding:0;}
  body{background:var(--bg);color:var(--text);font-family:-apple-system,BlinkMacSystemFont,"Segoe UI",sans-serif;font-size:13px;height:100vh;display:flex;flex-direction:column;overflow:hidden;}
  header{background:var(--surface);border-bottom:1px solid var(--border);padding:8px 18px;display:flex;align-items:center;gap:12px;flex-shrink:0;}
  header h1{font-size:15px;font-weight:700;}
  .health-badge{padding:3px 10px;border-radius:999px;font-size:12px;font-weight:600;background:#3b1a1a;}
  header .gen{font-size:11px;color:var(--text2);margin-left:auto;}
  .kpi-strip{display:flex;gap:8px;padding:8px 18px;flex-shrink:0;overflow-x:auto;}
  .kpi{background:var(--surface);border:1px solid var(--border);border-radius:8px;padding:7px 14px;white-space:nowrap;display:flex;align-items:baseline;gap:8px;}
  .kpi .val{font-size:20px;font-weight:700;line-height:1;}
  .kpi .lbl{font-size:10px;color:var(--text2);text-transform:uppercase;letter-spacing:.4px;}
  .kpi.red .val{color:var(--red);}.kpi.amber .val{color:var(--amber);}.kpi.green .val{color:var(--green);}.kpi.blue .val{color:var(--blue);}
  .tab-bar{display:flex;padding:0 18px;flex-shrink:0;border-bottom:1px solid var(--border);}
  .tab-btn{padding:7px 18px;font-size:12px;font-weight:500;cursor:pointer;border:none;background:transparent;color:var(--text2);border-bottom:2px solid transparent;margin-bottom:-1px;transition:color .12s,border-color .12s;}
  .tab-btn:hover{color:var(--text);} .tab-btn.active{color:var(--blue);border-bottom-color:var(--blue);}
  .charts-strip{display:grid;grid-template-columns:220px 1fr 200px;gap:8px;padding:8px 18px;flex-shrink:0;height:165px;}
  .panel{background:var(--surface);border:1px solid var(--border);border-radius:8px;padding:10px;display:flex;flex-direction:column;overflow:hidden;}
  .panel-title{font-size:10px;text-transform:uppercase;letter-spacing:.5px;color:var(--text2);margin-bottom:6px;flex-shrink:0;}
  .chart-wrap{flex:1;position:relative;min-height:0;}
  .type-list{display:flex;flex-direction:column;gap:4px;overflow-y:auto;flex:1;}
  .type-tile{display:flex;align-items:center;justify-content:space-between;background:var(--surface2);border:1px solid var(--border);border-radius:6px;padding:5px 10px;cursor:pointer;transition:border-color .12s,background .12s;user-select:none;}
  .type-tile:hover{border-color:var(--accent);} .type-tile.active{border-color:var(--blue);background:#1a2744;}
  .type-tile .type-label{font-size:12px;} .type-tile .type-badge{background:var(--accent);color:#fff;border-radius:999px;padding:1px 7px;font-size:11px;font-weight:600;}
  .main-area{flex:1;display:flex;overflow:hidden;padding:0 18px 10px;gap:8px;}
  .task-section{flex:1;display:flex;flex-direction:column;overflow:hidden;}
  .section-bar{display:flex;align-items:center;gap:10px;padding:6px 0 5px;flex-shrink:0;}
  .section-bar h3{font-size:12px;font-weight:600;}
  .filter-info{font-size:11px;color:var(--text2);}
  .filter-clear{font-size:11px;color:var(--blue);cursor:pointer;text-decoration:underline;}
  .tbl-wrap{flex:1;overflow-y:auto;border:1px solid var(--border);border-radius:8px;}
  table{width:100%;border-collapse:collapse;}
  thead th{position:sticky;top:0;background:var(--surface2);padding:6px 8px;text-align:left;font-size:10px;text-transform:uppercase;letter-spacing:.4px;color:var(--text2);border-bottom:1px solid var(--border);white-space:nowrap;}
  tbody tr{border-bottom:1px solid var(--border);cursor:pointer;}
  tbody tr:hover{background:var(--surface2);}
  tbody tr.selected{background:#1a2744;border-left:2px solid var(--blue);}
  tbody td{padding:5px 8px;font-size:12px;vertical-align:top;}
  .pill{display:inline-block;padding:1px 7px;border-radius:999px;font-size:10px;font-weight:500;margin:1px 2px;}
  .pill.risk{background:#3b1a1a;color:#ef9a9a;} .pill.issue{background:#1a2a3b;color:#90caf9;}
  .pill.slip{background:#3b2600;color:#ffb74d;} .pill.done{background:#1a2e1a;color:#a5d6a7;}
  .t-name{font-weight:500;max-width:240px;} .t-ws{font-size:10px;color:var(--text2);margin-top:1px;}
  .empty{text-align:center;padding:30px;color:var(--text2);}
  .dep-panel{width:0;flex-shrink:0;background:var(--surface);border:1px solid var(--border);border-radius:8px;overflow:hidden;transition:width .2s ease;display:flex;flex-direction:column;}
  .dep-panel.open{width:420px;}
  .dep-header{padding:10px 14px;border-bottom:1px solid var(--border);display:flex;align-items:center;justify-content:space-between;flex-shrink:0;}
  .dep-header h4{font-size:12px;font-weight:600;}
  .dep-close{cursor:pointer;color:var(--text2);font-size:16px;line-height:1;padding:0 4px;}
  .dep-close:hover{color:var(--text);}
  .dep-body{flex:1;overflow-y:auto;padding:14px;}
  .dep-flow{display:flex;align-items:flex-start;gap:0;margin-bottom:16px;}
  .dep-col{display:flex;flex-direction:column;gap:6px;min-width:110px;} .dep-col.center{min-width:130px;}
  .dep-arrows{display:flex;flex-direction:column;justify-content:center;align-items:center;padding:0 6px;gap:6px;}
  .dep-box{border-radius:6px;padding:6px 8px;font-size:11px;line-height:1.3;}
  .dep-box.pred{background:#1e2a1e;border:1px solid #2e4a2e;color:#a5d6a7;} .dep-box.pred.slipped{background:#2a1e1e;border-color:#4a2e2e;color:#ef9a9a;}
  .dep-box.focal{background:#1a2744;border:2px solid var(--blue);color:var(--text);font-weight:600;} .dep-box.focal.slipped{background:#2a1f0e;border-color:#ff9800;color:#ffb74d;}
  .dep-box.suc{background:#1e1e2a;border:1px solid #2e2e4a;color:#90caf9;} .dep-box.suc.slipped{background:#2a1e1e;border-color:#4a2e2e;color:#ef9a9a;}
  .dep-box .box-name{font-weight:500;margin-bottom:2px;} .dep-box .box-meta{font-size:10px;color:var(--text2);} .dep-box .box-slip{font-size:10px;color:#ffb74d;margin-top:2px;}
  .arr{font-size:18px;color:var(--text2);line-height:1;}
  .dep-section{margin-bottom:14px;} .dep-section h5{font-size:10px;text-transform:uppercase;letter-spacing:.5px;color:var(--text2);margin-bottom:6px;}
  .dep-detail-row{background:var(--surface2);border-radius:6px;padding:7px 10px;margin-bottom:5px;font-size:12px;}
  .dep-detail-row .dr-label{color:var(--text2);font-size:10px;margin-bottom:2px;}
  /* gantt */
  .gantt-outer{flex:1;display:flex;flex-direction:column;padding:0 18px 10px;overflow:hidden;}
  .gantt-controls{display:flex;align-items:center;gap:10px;padding:8px 0 6px;flex-shrink:0;flex-wrap:wrap;}
  .gantt-controls select{background:var(--surface2);color:var(--text);border:1px solid var(--border);border-radius:6px;padding:4px 8px;font-size:12px;}
  .gantt-controls label{font-size:12px;color:var(--text2);display:flex;align-items:center;gap:5px;cursor:pointer;}
  .gantt-controls input[type=checkbox]{accent-color:var(--blue);}
  .gantt-legend{display:flex;gap:12px;font-size:11px;color:var(--text2);margin-left:auto;flex-wrap:wrap;}
  .gantt-legend span{display:flex;align-items:center;gap:4px;}
  .gantt-count{font-size:11px;color:var(--text2);}
  .gantt-body{flex:1;display:flex;overflow:hidden;border:1px solid var(--border);border-radius:8px;}
  .gantt-names{width:220px;flex-shrink:0;overflow:hidden;border-right:1px solid var(--border);display:flex;flex-direction:column;}
  .gantt-name-hdr{height:36px;background:var(--surface2);border-bottom:1px solid var(--border);display:flex;align-items:center;padding:0 10px;font-size:10px;text-transform:uppercase;letter-spacing:.5px;color:var(--text2);flex-shrink:0;}
  .gantt-name-list{flex:1;overflow-y:hidden;}
  .gantt-chart-wrap{flex:1;overflow:auto;}
  #ganttTip{position:fixed;background:#1a1d27;border:1px solid #2e3350;border-radius:6px;padding:8px 12px;font-size:11px;pointer-events:none;display:none;z-index:9999;max-width:320px;line-height:1.6;white-space:pre-wrap;}
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
</div>

<!-- TASKS TAB -->
<div id="tasksContent" style="display:contents">
  <div class="charts-strip">
    <div class="panel">
      <div class="panel-title">Tasks by Status</div>
      <div class="chart-wrap"><canvas id="statusChart"></canvas></div>
    </div>
    <div class="panel">
      <div class="panel-title">Schedule Pressure by Workstream</div>
      <div class="chart-wrap"><canvas id="wsChart"></canvas></div>
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
            <th>ID</th><th>Task Name</th><th>Status</th><th>Start</th><th>Finish</th>
            <th>Baseline</th><th>Slip</th><th>↑Pred</th><th>↓Suc</th><th>Linked Issues / Risks</th>
          </tr></thead>
          <tbody id="taskBody"></tbody>
        </table>
      </div>
    </div>
    <div class="dep-panel" id="depPanel">
      <div class="dep-header">
        <h4 id="depTitle">Dependency Chain</h4>
        <span class="dep-close" onclick="closeDepPanel()">✕</span>
      </div>
      <div class="dep-body" id="depBody"></div>
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
      <label><input type="checkbox" id="ganttMilestones" checked onchange="ganttRendered=false;renderGantt()"> Milestones</label>
      <label><input type="checkbox" id="ganttIssues" checked onchange="ganttRendered=false;renderGantt()"> Has issues</label>
      <label><input type="checkbox" id="ganttSlipped" checked onchange="ganttRendered=false;renderGantt()"> Slipped</label>
      <span class="gantt-count" id="ganttCount"></span>
      <div class="gantt-legend">
        <span><svg width="12" height="10"><polygon points="6,0 12,5 6,10 0,5" fill="#22c55e"/></svg> Milestone</span>
        <span><svg width="12" height="10"><polygon points="6,0 12,5 6,10 0,5" fill="#ef5350"/></svg> Slipped milestone</span>
        <span><svg width="12" height="10"><rect width="12" height="10" rx="2" fill="#92400e"/></svg> Slipped task</span>
        <span><svg width="12" height="10"><rect width="12" height="10" rx="2" fill="#1e3a8a"/></svg> In progress</span>
        <span><svg width="10" height="10"><circle cx="5" cy="5" r="5" fill="#ef5350"/></svg> Has issue</span>
      </div>
    </div>
    <div class="gantt-body">
      <div class="gantt-names">
        <div class="gantt-name-hdr">Task / Milestone</div>
        <div class="gantt-name-list" id="ganttNames"></div>
      </div>
      <div class="gantt-chart-wrap" id="ganttChartWrap">
        <div id="ganttSvgWrap"></div>
      </div>
    </div>
  </div>
</div>
<div id="ganttTip"></div>

<script>
const DATA = ${data};
let activeFilter = null;
let selectedTaskId = null;
let ganttRendered = false;
const taskById = new Map(DATA.tasks.map(t => [t.id, t]));
const fmt = d => d ? new Date(d).toLocaleDateString("en-US",{month:"short",day:"numeric",year:"2-digit"}) : "—";
const esc = s => String(s||"").replace(/&/g,"&amp;").replace(/</g,"&lt;").replace(/>/g,"&gt;");
const trunc = (s,n) => s && s.length>n ? s.slice(0,n)+"…" : (s||"");

// ── Tab switching ──────────────────────────────────────────────────────────
function switchTab(tab) {
  const isGantt = tab === "gantt";
  document.getElementById("tabTasks").classList.toggle("active", !isGantt);
  document.getElementById("tabGantt").classList.toggle("active", isGantt);
  const tc = document.getElementById("tasksContent");
  tc.style.display = isGantt ? "none" : "contents";
  const gc = document.getElementById("ganttContent");
  gc.style.display = isGantt ? "flex" : "none";
  if (isGantt && !ganttRendered) { renderGantt(); ganttRendered = true; }
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
  const statusColors = DATA.statusBreakdown.map((_,i) => \`hsl(\${(i*47+200)%360},60%,55%)\`);
  new Chart(document.getElementById("statusChart"),{type:"doughnut",data:{labels:DATA.statusBreakdown.map(x=>x.s),datasets:[{data:DATA.statusBreakdown.map(x=>x.c),backgroundColor:statusColors,borderWidth:0}]},options:{responsive:true,maintainAspectRatio:false,plugins:{legend:{position:"bottom",labels:{color:"#9096b8",font:{size:10},boxWidth:10,padding:6}}}}});
  const ws = DATA.topWorkstreams.slice(0,8);
  new Chart(document.getElementById("wsChart"),{type:"bar",data:{labels:ws.map(x=>trunc(x.ws,20)),datasets:[{data:ws.map(x=>x.cnt),backgroundColor:"#5c6bc0",borderRadius:3}]},options:{indexAxis:"y",responsive:true,maintainAspectRatio:false,plugins:{legend:{display:false}},scales:{x:{ticks:{color:"#9096b8",font:{size:10}},grid:{color:"#2e3350"}},y:{ticks:{color:"#e8eaf6",font:{size:10}},grid:{display:false}}}}});
  const typeList = document.getElementById("typeList");
  DATA.riskTypeBreakdown.forEach(rt => {
    const tile = document.createElement("div");
    tile.className = "type-tile"; tile.dataset.typeKey = rt.typeKey;
    tile.innerHTML = \`<span class="type-label">\${esc(rt.typeKey)}</span><span class="type-badge">\${rt.count}</span>\`;
    tile.addEventListener("click", () => toggleFilter(rt.typeKey, tile));
    typeList.appendChild(tile);
  });
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
  if (!tasks.length) { tbody.innerHTML = '<tr><td colspan="10" class="empty">No matching tasks.</td></tr>'; return; }
  tbody.innerHTML = tasks.slice(0,400).map(t => {
    const pills = t.linkedRisks.map(r => \`<span class="pill \${/issue/i.test(r.type)?"issue":"risk"}" title="\${esc(r.severity)}">\${esc(trunc(r.name,26))}</span>\`).join("");
    const slipCell = t.slipDays != null ? \`<span class="pill slip">+\${t.slipDays}d</span>\` : t.pct >= 100 ? \`<span class="pill done">Done</span>\` : "—";
    return \`<tr class="\${selectedTaskId===t.id?" selected":""}" onclick="selectTask(event,'\${t.id}')">
      <td style="color:var(--text2);white-space:nowrap">\${t.uid||"—"}</td>
      <td><div class="t-name">\${esc(trunc(t.name,55))}</div><div class="t-ws">\${esc(t.workstream)}</div></td>
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
  selectedTaskId = taskId;
  document.querySelectorAll("#taskBody tr").forEach(r => r.classList.remove("selected"));
  evt.currentTarget.classList.add("selected");
  openDepPanel(taskId);
}
function openDepPanel(taskId) {
  const task = taskById.get(taskId); if (!task) return;
  document.getElementById("depTitle").textContent = trunc(task.name, 38);
  document.getElementById("depPanel").classList.add("open");
  renderDepPanel(task);
}
function closeDepPanel() {
  document.getElementById("depPanel").classList.remove("open");
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
  const wsFilter  = document.getElementById("ganttWs").value;
  const showMiles = document.getElementById("ganttMilestones").checked;
  const showIssue = document.getElementById("ganttIssues").checked;
  const showSlip  = document.getElementById("ganttSlipped").checked;

  let tasks = DATA.tasks.filter(t => {
    if (t.pct >= 100) return false;
    if (wsFilter && t.workstream !== wsFilter) return false;
    return (showMiles && t.milestone) ||
           (showIssue && t.linkedRisks.length > 0) ||
           (showSlip  && t.isSlipped);
  });
  tasks.sort((a,b) => {
    const am = a.milestone ? 0 : 1, bm = b.milestone ? 0 : 1;
    if (am !== bm) return am - bm;
    return (a.finish||"").localeCompare(b.finish||"");
  });

  document.getElementById("ganttCount").textContent = \`\${tasks.length} items\`;
  if (!tasks.length) {
    document.getElementById("ganttNames").innerHTML = '<div style="padding:20px;color:var(--text2);font-size:12px">No tasks match filters.</div>';
    document.getElementById("ganttSvgWrap").innerHTML = "";
    return;
  }

  const allMs = tasks.flatMap(t => [t.start,t.finish,t.baselineFinish].filter(Boolean).map(d=>new Date(d).getTime())).filter(n=>!isNaN(n));
  const minDate = new Date(Math.min(...allMs)); minDate.setDate(1);
  const maxDate = new Date(Math.max(...allMs)); maxDate.setMonth(maxDate.getMonth()+1,1);
  const totalDays = Math.max(1, (maxDate - minDate)/86400000);
  const ppd = Math.max(1.2, Math.min(5, 1100/totalDays));
  const ROW_H=28, HDR_H=36;
  const svgW = Math.ceil(totalDays*ppd)+30;
  const svgH = HDR_H + tasks.length*ROW_H;
  const toX = d => { if (!d) return null; const ms=new Date(d).getTime(); return isNaN(ms)?null:Math.round((ms-minDate.getTime())/86400000*ppd)+10; };
  const todayX = toX(new Date().toISOString().slice(0,10));

  // Month gridlines
  const months=[];
  const cur=new Date(minDate);
  while(cur<maxDate){ months.push({x:toX(cur.toISOString().slice(0,10)),lbl:cur.toLocaleDateString("en-US",{month:"short",year:"2-digit"})}); cur.setMonth(cur.getMonth()+1); }

  const svg=[
    \`<svg xmlns="http://www.w3.org/2000/svg" width="\${svgW}" height="\${svgH}" style="display:block;min-width:\${svgW}px">\`,
    \`<rect width="\${svgW}" height="\${svgH}" fill="#0f1117"/>\`,
  ];
  months.forEach(m=>{
    svg.push(\`<line x1="\${m.x}" y1="0" x2="\${m.x}" y2="\${svgH}" stroke="#2e3350" stroke-width="1"/>\`);
    svg.push(\`<text x="\${m.x+4}" y="14" fill="#9096b8" font-size="10" font-family="sans-serif">\${m.lbl}</text>\`);
  });
  svg.push(\`<rect width="\${svgW}" height="\${HDR_H}" fill="#1a1d27"/>\`);
  svg.push(\`<line x1="0" y1="\${HDR_H}" x2="\${svgW}" y2="\${HDR_H}" stroke="#2e3350"/>\`);
  if(todayX!=null){
    svg.push(\`<line x1="\${todayX}" y1="\${HDR_H}" x2="\${todayX}" y2="\${svgH}" stroke="#ef5350" stroke-width="1.5" stroke-dasharray="4,3" opacity="0.7"/>\`);
    svg.push(\`<text x="\${todayX+3}" y="\${HDR_H-5}" fill="#ef5350" font-size="9" font-family="sans-serif">Today</text>\`);
  }

  const nameRows=[];
  tasks.forEach((t,i)=>{
    const y=HDR_H+i*ROW_H, rowBg=i%2===0?"#1a1d27":"#0f1117";
    svg.push(\`<rect x="0" y="\${y}" width="\${svgW}" height="\${ROW_H}" fill="\${rowBg}"/>\`);
    const x1=toX(t.start), x2=toX(t.finish), xBL=t.baselineFinish?toX(t.baselineFinish):null;
    const my=y+ROW_H/2;

    // Tooltip lines
    const tip=[
      \`#\${t.uid||"?"} \${t.name}\`,
      \`Workstream: \${t.workstream}\`,
      \`Finish: \${fmt(t.finish)}\${t.baselineFinish?" (baseline: "+fmt(t.baselineFinish)+")":""}\`,
      t.slipDays?\`⚠ Slipped +\${t.slipDays} days\`:"",
      ...t.linkedRisks.map(r=>\`⚑ \${r.type}: \${r.name} [\${r.severity}] — \${r.category}\`),
    ].filter(Boolean).join("\\n");
    const tipSvg=tip.replace(/&/g,"&amp;").replace(/</g,"&lt;").replace(/>/g,"&gt;");

    if(!t.milestone){
      const barColor=t.isSlipped?"#92400e":t.pct>0?"#1e3a8a":"#374151";
      const bdrColor=t.isSlipped?"#f59e0b":t.pct>0?"#3b82f6":"#6b7280";
      if(x1!=null&&x2!=null){
        const bw=Math.max(6,x2-x1);
        svg.push(\`<rect x="\${x1}" y="\${y+7}" width="\${bw}" height="\${ROW_H-14}" rx="3" fill="\${barColor}" stroke="\${bdrColor}" stroke-width="0.5"><title>\${tipSvg}</title></rect>\`);
        if(t.pct>0&&t.pct<100) svg.push(\`<rect x="\${x1}" y="\${y+7}" width="\${Math.max(4,Math.round(bw*t.pct/100))}" height="\${ROW_H-14}" rx="3" fill="\${bdrColor}" opacity="0.9"/>\`);
        if(xBL&&Math.abs(xBL-x2)>3) svg.push(\`<line x1="\${xBL}" y1="\${y+4}" x2="\${xBL}" y2="\${y+ROW_H-4}" stroke="#ff9800" stroke-width="2.5" opacity="0.7"><title>Baseline: \${fmt(t.baselineFinish)}</title></line>\`);
        if(t.linkedRisks.length>0){
          const dx=x2+7, it=t.linkedRisks.map(r=>r.type+": "+r.name+" ["+r.severity+"] — "+r.category).join("; ").replace(/&/g,"&amp;").replace(/</g,"&lt;");
          svg.push(\`<circle cx="\${dx}" cy="\${my}" r="6" fill="#ef5350" stroke="#450a0a" stroke-width="1"><title>\${it}</title></circle>\`);
          svg.push(\`<text x="\${dx}" y="\${my+4}" text-anchor="middle" font-size="9" fill="white" font-weight="bold" font-family="sans-serif">!</text>\`);
        }
      }
    } else {
      const mx=x2??x1;
      if(mx!=null){
        const d=9, mc=t.isSlipped?"#ef5350":"#22c55e", ms=t.isSlipped?"#7f1d1d":"#14532d";
        svg.push(\`<polygon points="\${mx},\${my-d} \${mx+d},\${my} \${mx},\${my+d} \${mx-d},\${my}" fill="\${mc}" stroke="\${ms}" stroke-width="1.5"><title>\${tipSvg}</title></polygon>\`);
        svg.push(\`<text x="\${mx+d+5}" y="\${my+4}" fill="\${mc}" font-size="10" font-family="sans-serif" font-weight="600">\${esc(trunc(t.name,28))}</text>\`);
        if(xBL&&Math.abs(xBL-mx)>3) svg.push(\`<line x1="\${xBL}" y1="\${my}" x2="\${mx}" y2="\${my}" stroke="#ff9800" stroke-width="1.5" stroke-dasharray="3,2" opacity="0.7"><title>Baseline: \${fmt(t.baselineFinish)}</title></line>\`);
        if(t.linkedRisks.length>0){
          const it=t.linkedRisks.map(r=>r.type+": "+r.name+" ["+r.severity+"] — "+r.category).join("; ").replace(/&/g,"&amp;").replace(/</g,"&lt;");
          svg.push(\`<circle cx="\${mx}" cy="\${my-d-8}" r="6" fill="#ef5350" stroke="#450a0a" stroke-width="1"><title>\${it}</title></circle>\`);
          svg.push(\`<text x="\${mx}" y="\${my-d-4}" text-anchor="middle" font-size="9" fill="white" font-weight="bold" font-family="sans-serif">!</text>\`);
        }
      }
    }

    const nc=t.milestone?"#fbbf24":t.isSlipped?"#fca5a5":"#e8eaf6";
    const ic=t.milestone?"🔷":t.linkedRisks.length>0?"⚠":"·";
    nameRows.push(\`<div style="height:\${ROW_H}px;line-height:\${ROW_H}px;padding:0 8px;font-size:11px;overflow:hidden;text-overflow:ellipsis;white-space:nowrap;background:\${rowBg};color:\${nc};border-bottom:1px solid #2e3350" title="\${esc(t.name)} — \${esc(t.workstream)}">\${ic} \${esc(trunc(t.name,24))}</div>\`);
  });
  svg.push("</svg>");

  document.getElementById("ganttNames").innerHTML = nameRows.join("");
  document.getElementById("ganttSvgWrap").innerHTML = svg.join("");

  // Sync vertical scroll
  const cw=document.getElementById("ganttChartWrap"), nl=document.getElementById("ganttNames");
  cw.onscroll=()=>{ nl.scrollTop=cw.scrollTop; };
  nl.onscroll=()=>{ cw.scrollTop=nl.scrollTop; };

  // Styled tooltip
  const tip2=document.getElementById("ganttTip");
  const svgEl=document.getElementById("ganttSvgWrap").querySelector("svg");
  if(svgEl){
    svgEl.addEventListener("mousemove",e=>{
      const el=e.target.closest("[title]");
      if(el&&el.getAttribute("title")){
        tip2.style.display="block";
        tip2.style.left=(e.clientX+14)+"px";
        tip2.style.top=(e.clientY-10)+"px";
        tip2.innerHTML=el.getAttribute("title").replace(/&amp;/g,"&").replace(/&lt;/g,"<").replace(/&gt;/g,">").replace(/\\n/g,"<br>").replace(/⚑/g,"<span style='color:#ef5350'>⚑</span>").replace(/⚠/g,"<span style='color:#f59e0b'>⚠</span>");
      } else { tip2.style.display="none"; }
    });
    svgEl.addEventListener("mouseleave",()=>{ tip2.style.display="none"; });
  }

  // Scroll to today
  setTimeout(()=>{
    const cw2=document.getElementById("ganttChartWrap"); if(!cw2) return;
    const line=cw2.querySelector("line[stroke='#ef5350']");
    if(line){ const x=parseFloat(line.getAttribute("x1")||0); cw2.scrollLeft=Math.max(0,x-300); }
  },80);
}

init();
<\/script>
</body>
</html>`;
}

main().catch(e => { console.error(e); process.exit(1); });

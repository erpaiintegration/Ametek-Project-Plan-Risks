/* eslint-disable no-console */
require("dotenv/config");
const { Client } = require("@notionhq/client");
const fs = require("fs");
const path = require("path");

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
function relCountByNames(props, names) {
  for (const n of names) {
    const p = props[n];
    if (p?.type === "relation") return (p.relation || []).length;
  }
  return 0;
}
function daysDiff(a, b) {
  if (!a || !b) return null;
  const da = new Date(a), db = new Date(b);
  if (isNaN(da) || isNaN(db)) return null;
  return Math.round((da - db) / 86400000);
}

async function fetchAll(dbId, label) {
  const out = [];
  let cursor;
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

  const riskRecords = [];
  const taskIdToRisks = new Map();

  for (const r of rawRisks) {
    const p = r.properties || {};
    const status = selVal(p.Status) || "";
    if (/resolved|closed/i.test(status)) continue;

    const type = selVal(p.Type) || "Risk";
    const category = selVal(p["Risk Type"]) || selVal(p.Category) || selVal(p.Family) || "(Uncategorized)";
    const severity = selVal(p.Severity) || selVal(p.Priority) || selVal(p.Impact) || "(Unrated)";
    const titleObj = Object.values(p).find(x => x?.type === "title");
    const name = titleObj ? text(titleObj.title || []) : "(Untitled)";

    const linkedTaskIds = [
      ...relIds(p.Tasks), ...relIds(p["Related Tasks"]), ...relIds(p["Impacted Tasks"]), ...relIds(p.Task), ...relIds(p["Task Links"])
    ];

    const rec = { id: r.id, name, type, category, severity, status, linkedTaskIds, typeKey: `${type} — ${category}` };
    riskRecords.push(rec);

    for (const tid of linkedTaskIds) {
      if (!taskIdToRisks.has(tid)) taskIdToRisks.set(tid, []);
      taskIdToRisks.get(tid).push({ name, type, severity, category, status });
    }
  }

  const riskTypeMap = new Map();
  for (const rec of riskRecords) riskTypeMap.set(rec.typeKey, (riskTypeMap.get(rec.typeKey) || 0) + 1);
  const riskTypeBreakdown = [...riskTypeMap.entries()].sort((a, b) => b[1] - a[1]).map(([typeKey, count]) => ({ typeKey, count }));

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
      id: t.id, uid, name, workstream, start, finish, baselineFinish, pct, milestone,
      predCount, sucCount, predIds, sucIds,
      slipDays: isSlipped ? slipDays : null,
      isSlipped, isOverdueStart, isDue14, status: statusRaw, linkedRisks
    });
  }

  const total = tasks.length;
  const slipRatePct = totalOpen ? ((slippedOpen / totalOpen) * 100).toFixed(1) : "0.0";
  const openIssues = riskRecords.filter(r => /issue/i.test(r.type));
  const openRisks = riskRecords.filter(r => !/issue/i.test(r.type));
  const topWorkstreams = [...workstreamPressure.entries()].sort((a, b) => b[1] - a[1]).slice(0, 8).map(([ws, cnt]) => ({ ws, cnt }));
  const statusBreakdown = [...statusCounts.entries()].sort((a, b) => b[1] - a[1]).map(([s, c]) => ({ s, c }));

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

  fs.writeFileSync(OUT_FILE, buildHtml(payload), "utf8");
  console.log(`\nDashboard written to: ${OUT_FILE}`);
}

function buildHtml(payload) {
  const data = JSON.stringify(payload).replace(/<\/script>/gi, "<\\/script>");
  const workstreams = [...new Set(payload.tasks.map(t => t.workstream))].sort();
  const wsOptions = workstreams.map(w => `<option value="${w.replace(/"/g, "&quot;")}">${w}</option>`).join("\n");
  const chartScript = CHARTJS_INLINE
    ? `<script>${CHARTJS_INLINE.replace(/<\/script>/gi, "<\\/script>")}<\/script>`
    : `<script src="https://cdn.jsdelivr.net/npm/chart.js@4.4.3/dist/chart.umd.min.js"><\/script>`;

  return `<!doctype html><html><head><meta charset="UTF-8"><meta name="viewport" content="width=device-width,initial-scale=1"><title>Ametek Dashboard</title>${chartScript}<style>
body{margin:0;font-family:Segoe UI,Arial,sans-serif;background:#0f1117;color:#e8eaf6}header{padding:8px 14px;background:#1a1d27;border-bottom:1px solid #2e3350;display:flex;gap:10px;align-items:center}.gen{margin-left:auto;color:#9096b8;font-size:11px}.tabs{display:flex;padding:0 14px;border-bottom:1px solid #2e3350}.tab{background:transparent;border:none;color:#9ca3af;padding:8px 14px;cursor:pointer}.tab.active{color:#60a5fa;border-bottom:2px solid #60a5fa}.kpis{display:flex;gap:8px;padding:8px 14px;overflow:auto}.kpi{background:#1a1d27;border:1px solid #2e3350;border-radius:8px;padding:6px 10px;white-space:nowrap}.kpi b{font-size:18px}.layout{display:grid;grid-template-columns:1fr 380px;gap:8px;padding:8px 14px;height:52vh}.panel{background:#1a1d27;border:1px solid #2e3350;border-radius:8px;overflow:hidden}.panel h3{margin:0;padding:8px 10px;font-size:12px;color:#cbd5e1;border-bottom:1px solid #2e3350}.tbl{height:calc(100% - 34px);overflow:auto}table{width:100%;border-collapse:collapse}th,td{padding:6px 8px;border-bottom:1px solid #2e3350;font-size:12px;text-align:left}th{position:sticky;top:0;background:#22263a;color:#9ca3af;font-size:10px;text-transform:uppercase}
#ganttWrap{display:none;padding:8px 14px 10px;height:calc(100vh - 120px)}.gctrl{display:flex;gap:8px;align-items:center;flex-wrap:wrap;margin-bottom:8px}.gctrl select{padding:4px 6px;border:1px solid #cbd5e1;border-radius:6px;background:#fff;color:#111827;font-size:12px}.gctrl label{font-size:12px;color:#334155;display:flex;gap:4px;align-items:center}.gbox{height:calc(100% - 40px);display:grid;grid-template-columns:300px 1fr;border:1px solid #d1d5db;border-radius:8px;overflow:hidden;background:#fff}.gleft{border-right:1px solid #d1d5db;display:flex;flex-direction:column}.ghdr{height:52px;background:#f8fafc;border-bottom:1px solid #d1d5db;padding:8px 10px;font-size:10px;color:#64748b;text-transform:uppercase}.grows{flex:1;overflow:hidden}.gright{overflow:auto}#gsvg{min-height:100%}#gTip{position:fixed;display:none;pointer-events:none;z-index:9999;background:#fff;color:#111827;border:1px solid #d1d5db;border-radius:6px;padding:8px 10px;font-size:11px;max-width:360px;line-height:1.4;box-shadow:0 10px 22px rgba(0,0,0,.2)}.depmini{font-size:10px;color:#475569;border:1px solid #cbd5e1;background:#f8fafc;border-radius:999px;padding:1px 6px;margin-left:6px}
</style></head><body>
<header><b>📊 Ametek SAP S4 Dashboard</b><span class="gen" id="gen"></span></header>
<div class="kpis" id="kpis"></div>
<div class="tabs"><button class="tab active" id="tabT">📋 Tasks</button><button class="tab" id="tabG">📅 Gantt</button></div>
<div class="layout" id="tasksWrap"><div class="panel"><h3>Tasks at risk</h3><div class="tbl"><table><thead><tr><th>ID</th><th>Task</th><th>Status</th><th>Start</th><th>Finish</th><th>Slip</th><th>Pred</th><th>Suc</th></tr></thead><tbody id="taskBody"></tbody></table></div></div><div class="panel"><h3>Dependency details</h3><div class="tbl" id="depBody" style="padding:10px;color:#9ca3af">Select a task row</div></div></div>
<div id="ganttWrap"><div class="gctrl"><label>Workstream <select id="gWs"><option value="">All</option>${wsOptions}</select></label><label>Risk Type <select id="gRisk"><option value="">All</option><option value="__ANY__">Any linked risk/issue</option></select></label><label>Window <select id="gWin"><option value="all">All</option><option value="365">12 mo</option><option value="180" selected>6 mo</option><option value="90">3 mo</option></select></label><label>Zoom <select id="gZoom"><option value="0.8">Compact</option><option value="1" selected>Normal</option><option value="1.4">Detailed</option></select></label><label><input type="checkbox" id="gMil" checked> Milestones</label><label><input type="checkbox" id="gIssue" checked> Has issues</label><label><input type="checkbox" id="gSlip" checked> Slipped</label><span id="gCnt" style="color:#64748b;font-size:12px"></span></div><div class="gbox"><div class="gleft"><div class="ghdr">Task / Milestone</div><div class="grows" id="grows"></div></div><div class="gright" id="gright"><div id="gsvg"></div></div></div></div>
<div id="gTip"></div>
<script>
const DATA=${data};
const taskById=new Map(DATA.tasks.map(t=>[t.id,t]));
const fmt=d=>d?new Date(d).toLocaleDateString("en-US",{month:"short",day:"numeric",year:"2-digit"}):"—";
const esc=s=>String(s||"").replace(/&/g,"&amp;").replace(/</g,"&lt;").replace(/>/g,"&gt;");
const trunc=(s,n)=>s&&s.length>n?s.slice(0,n)+"…":(s||"");
const riskIcon=r=>{const sev=(r.severity||"").toLowerCase();const typ=(r.type||"").toLowerCase();if(sev.includes("critical")||sev.includes("high"))return"🔴";if(sev.includes("medium"))return"🟠";if(typ.includes("issue"))return"⚠️";return"🟡";};

document.getElementById("gen").textContent="Generated "+new Date(DATA.generatedAt).toLocaleString();
[{label:"Total",v:DATA.metrics.total},{label:"Open",v:DATA.metrics.totalOpen},{label:"Slipped",v:DATA.metrics.slippedOpen},{label:"Open Risks",v:DATA.metrics.openRisks},{label:"Open Issues",v:DATA.metrics.openIssues},{label:"Slip Rate",v:DATA.metrics.slipRatePct+"%"}].forEach(k=>{const d=document.createElement("div");d.className="kpi";d.innerHTML=`<b>${k.v}</b><div style=\"font-size:10px;color:#94a3b8\">${k.label}</div>`;document.getElementById("kpis").appendChild(d);});

function renderTasks(){
  const rows=DATA.tasks.filter(t=>t.isSlipped||t.isOverdueStart||t.isDue14||t.linkedRisks.length>0).sort((a,b)=>(b.slipDays||0)-(a.slipDays||0)).slice(0,400);
  const tb=document.getElementById("taskBody");
  tb.innerHTML=rows.map(t=>`<tr data-id="${t.id}"><td>${t.uid||"—"}</td><td>${esc(trunc(t.name,48))}<div style=\"font-size:10px;color:#9ca3af\">${esc(t.workstream)}</div></td><td>${esc(t.status)}</td><td>${fmt(t.start)}</td><td>${fmt(t.finish)}</td><td>${t.slipDays!=null?"+"+t.slipDays+"d":"—"}</td><td>${t.predCount||0}</td><td>${t.sucCount||0}</td></tr>`).join("");
  [...tb.querySelectorAll("tr")].forEach(r=>r.onclick=()=>{const t=taskById.get(r.dataset.id);const preds=(t.predIds||[]).map(id=>taskById.get(id)).filter(Boolean);const sucs=(t.sucIds||[]).map(id=>taskById.get(id)).filter(Boolean);document.getElementById("depBody").innerHTML=`<div style=\"padding:10px\"><div style=\"font-weight:600;margin-bottom:8px\">${esc(t.name)}</div><div style=\"font-size:12px;color:#94a3b8;margin-bottom:8px\">Start ${fmt(t.start)} → Finish ${fmt(t.finish)}</div><div style=\"margin-bottom:8px\">⬆ Pred ${preds.length} · ⬇ Succ ${sucs.length}</div><div style=\"font-size:12px\">${t.linkedRisks.map(x=>`<div>• ${riskIcon(x)} ${esc(x.type)} — ${esc(x.name)} (${esc(x.severity)})</div>`).join("")||"No linked risk/issue"}</div></div>`;});
}

function renderGantt(){
  const ws=document.getElementById("gWs").value,risk=document.getElementById("gRisk").value,win=document.getElementById("gWin").value,zoom=parseFloat(document.getElementById("gZoom").value||"1"),showMil=document.getElementById("gMil").checked,showIssue=document.getElementById("gIssue").checked,showSlip=document.getElementById("gSlip").checked;
  let rows=DATA.tasks.filter(t=>{if(t.pct>=100)return false;if(ws&&t.workstream!==ws)return false;if(risk==="__ANY__"&&t.linkedRisks.length===0)return false;if(risk&&risk!=="__ANY__"&&!t.linkedRisks.some(r=>(r.type+" — "+r.category)===risk))return false;return (showMil&&t.milestone)||(showIssue&&t.linkedRisks.length>0)||(showSlip&&t.isSlipped);});
  if(win!=="all"){const now=new Date(),d=Number(win),a=new Date(now.getTime()-d*0.2*86400000),b=new Date(now.getTime()+d*0.8*86400000);rows=rows.filter(t=>{const s=t.start?new Date(t.start):null,f=t.finish?new Date(t.finish):s;if(!s&&!f)return false;return (!f||f>=a)&&(!s||s<=b);});}
  rows.sort((a,b)=>(a.finish||"").localeCompare(b.finish||""));
  document.getElementById("gCnt").textContent=rows.length+" items";
  if(!rows.length){document.getElementById("grows").innerHTML='<div style="padding:12px;color:#64748b">No tasks match filters</div>';document.getElementById("gsvg").innerHTML="";return;}
  const all=rows.flatMap(t=>[t.start,t.finish,t.baselineFinish].filter(Boolean).map(d=>new Date(d).getTime())).filter(n=>!isNaN(n));
  const minDate=new Date(Math.min(...all));minDate.setDate(1);const maxDate=new Date(Math.max(...all));maxDate.setMonth(maxDate.getMonth()+1,1);
  const totalDays=Math.max(1,(maxDate-minDate)/86400000),ppd=Math.max(1.2,Math.min(8,(1100/totalDays)*zoom));
  const ROW=28,HDR=36,ML=44,norm=rows.filter(t=>!t.milestone),miles=rows.filter(t=>t.milestone),W=Math.ceil(totalDays*ppd)+30,H=HDR+ML+norm.length*ROW;
  const toX=d=>{if(!d)return null;const ms=new Date(d).getTime();return isNaN(ms)?null:Math.round((ms-minDate.getTime())/86400000*ppd)+10;},todayX=toX(new Date().toISOString().slice(0,10));
  const months=[];const c=new Date(minDate);while(c<maxDate){months.push({x:toX(c.toISOString().slice(0,10)),l:c.toLocaleDateString("en-US",{month:"short",year:"2-digit"})});c.setMonth(c.getMonth()+1);} 
  const svg=[`<svg xmlns=\"http://www.w3.org/2000/svg\" width=\"${W}\" height=\"${H}\" style=\"display:block;min-width:${W}px\">`,`<rect width=\"${W}\" height=\"${H}\" fill=\"#fff\"/>`];
  months.forEach(m=>{svg.push(`<line x1=\"${m.x}\" y1=\"0\" x2=\"${m.x}\" y2=\"${H}\" stroke=\"#e5e7eb\"/>`);svg.push(`<text x=\"${m.x+4}\" y=\"14\" fill=\"#475569\" font-size=\"10\">${m.l}</text>`);});
  svg.push(`<rect width=\"${W}\" height=\"${HDR}\" fill=\"#f8fafc\"/>`);svg.push(`<line x1=\"0\" y1=\"${HDR}\" x2=\"${W}\" y2=\"${HDR}\" stroke=\"#d1d5db\"/>`);svg.push(`<rect x=\"0\" y=\"${HDR}\" width=\"${W}\" height=\"${ML}\" fill=\"#f8fafc\"/>`);svg.push(`<line x1=\"0\" y1=\"${HDR+ML}\" x2=\"${W}\" y2=\"${HDR+ML}\" stroke=\"#d1d5db\"/>`);svg.push(`<text x=\"8\" y=\"${HDR+13}\" fill=\"#64748b\" font-size=\"10\">Milestones</text>`);
  if(todayX!=null){svg.push(`<line x1=\"${todayX}\" y1=\"${HDR}\" x2=\"${todayX}\" y2=\"${H}\" stroke=\"#ef5350\" stroke-width=\"1.5\" stroke-dasharray=\"4,3\"/>`);}
  const nameRows=[];
  norm.forEach((t,i)=>{const y=HDR+ML+i*ROW,rowBg=i%2===0?"#fff":"#f8fafc";svg.push(`<rect x=\"0\" y=\"${y}\" width=\"${W}\" height=\"${ROW}\" fill=\"${rowBg}\"/>`);const x1=toX(t.start),x2=toX(t.finish),xBL=t.baselineFinish?toX(t.baselineFinish):null,my=y+ROW/2,bar=t.isSlipped?"#f59e0b":t.pct>0?"#2563eb":"#94a3b8",bdr=t.isSlipped?"#b45309":t.pct>0?"#1d4ed8":"#64748b";
    const tip=[`#${t.uid||"?"} ${t.name}`,`Workstream: ${t.workstream}`,`Start/Finish: ${fmt(t.start)} → ${fmt(t.finish)}`,...t.linkedRisks.map(r=>`${riskIcon(r)} ${r.type}: ${r.name} [${r.severity}] — ${r.category}`)].join("\\n").replace(/&/g,"&amp;").replace(/</g,"&lt;").replace(/>/g,"&gt;");
    if(x1!=null&&x2!=null){const bw=Math.max(6,x2-x1);svg.push(`<rect x=\"${x1}\" y=\"${y+7}\" width=\"${bw}\" height=\"${ROW-14}\" rx=\"3\" fill=\"${bar}\" stroke=\"${bdr}\"><title>${tip}</title></rect>`);if(t.pct>0&&t.pct<100)svg.push(`<rect x=\"${x1}\" y=\"${y+7}\" width=\"${Math.max(4,Math.round(bw*t.pct/100))}\" height=\"${ROW-14}\" rx=\"3\" fill=\"${bdr}\" opacity=\"0.9\"/>`);if(xBL&&Math.abs(xBL-x2)>3)svg.push(`<line x1=\"${xBL}\" y1=\"${y+4}\" x2=\"${xBL}\" y2=\"${y+ROW-4}\" stroke=\"#f97316\" stroke-width=\"2\"/>`);const dt=fmt(t.start)+" → "+fmt(t.finish);if(bw>130)svg.push(`<text x=\"${x1+6}\" y=\"${my+3}\" fill=\"#fff\" font-size=\"9\" font-weight=\"600\">${dt}</text>`);else svg.push(`<text x=\"${x2+4}\" y=\"${my+3}\" fill=\"#334155\" font-size=\"9\">${dt}</text>`);if(t.linkedRisks.length){const icon=riskIcon(t.linkedRisks[0]);const rt=t.linkedRisks.map(r=>`${riskIcon(r)} ${r.type}: ${r.name} [${r.severity}] — ${r.category}`).join("; ").replace(/&/g,"&amp;").replace(/</g,"&lt;");const dx=x2-8;svg.push(`<circle cx=\"${dx}\" cy=\"${my}\" r=\"7\" fill=\"#fff\" stroke=\"#334155\"><title>${rt}</title></circle>`);svg.push(`<text x=\"${dx}\" y=\"${my+3}\" text-anchor=\"middle\" font-size=\"9\">${icon}</text>`);}if(risk){if((t.predIds||[]).length)svg.push(`<text x=\"${x1}\" y=\"${y+5}\" fill=\"#334155\" font-size=\"9\">⬆${(t.predIds||[]).length}</text>`);if((t.sucIds||[]).length)svg.push(`<text x=\"${x2+4}\" y=\"${y+5}\" fill=\"#334155\" font-size=\"9\">⬇${(t.sucIds||[]).length}</text>`);}}
    const depmini=risk?`<span class=\"depmini\">⬆${(t.predIds||[]).length} ⬇${(t.sucIds||[]).length}</span>`:"";nameRows.push(`<div style=\"height:${ROW}px;line-height:${ROW}px;padding:0 8px;font-size:11px;overflow:hidden;text-overflow:ellipsis;white-space:nowrap;background:${rowBg};color:#0f172a;border-bottom:1px solid #e5e7eb\" title=\"${esc(t.name)} — ${esc(t.workstream)}\">${t.linkedRisks.length?"⚠":"·"} ${esc(trunc(t.name,32))}${depmini}</div>`);
  });
  miles.forEach((t,idx)=>{const x1=toX(t.start),x2=toX(t.finish),mx=x2??x1;if(mx!=null){const my=HDR+24+(idx%2)*10,d=8,mc=t.isSlipped?"#ef4444":"#16a34a",ms=t.isSlipped?"#991b1b":"#166534";const tip=(t.name+"\\nMilestone date: "+fmt(t.finish||t.start)).replace(/&/g,"&amp;").replace(/</g,"&lt;").replace(/>/g,"&gt;");svg.push(`<polygon points=\"${mx},${my-d} ${mx+d},${my} ${mx},${my+d} ${mx-d},${my}\" fill=\"${mc}\" stroke=\"${ms}\"><title>${tip}</title></polygon>`);svg.push(`<text x=\"${mx+d+4}\" y=\"${my+3}\" fill=\"#334155\" font-size=\"9\" font-weight=\"600\">${esc(trunc(t.name,30))}</text>`);if(t.linkedRisks.length)svg.push(`<text x=\"${mx}\" y=\"${my-d-4}\" text-anchor=\"middle\" font-size=\"10\">${riskIcon(t.linkedRisks[0])}</text>`);}});
  svg.push("</svg>");document.getElementById("grows").innerHTML=nameRows.join("");document.getElementById("gsvg").innerHTML=svg.join("");const gr=document.getElementById("gright"),gl=document.getElementById("grows");gr.onscroll=()=>{gl.scrollTop=gr.scrollTop;};gl.onscroll=()=>{gr.scrollTop=gl.scrollTop;};
  const tip=document.getElementById("gTip"),svgEl=document.getElementById("gsvg").querySelector("svg");if(svgEl){svgEl.addEventListener("mousemove",e=>{const el=e.target.closest("[title]");if(el&&el.getAttribute("title")){tip.style.display="block";tip.style.left=(e.clientX+14)+"px";tip.style.top=(e.clientY-10)+"px";tip.innerHTML=el.getAttribute("title").replace(/&amp;/g,"&").replace(/&lt;/g,"<").replace(/&gt;/g,">").replace(/\\n/g,"<br>");}else tip.style.display="none";});svgEl.addEventListener("mouseleave",()=>{tip.style.display="none";});}
}

document.getElementById("tabT").onclick=()=>{document.getElementById("tabT").classList.add("active");document.getElementById("tabG").classList.remove("active");document.getElementById("tasksWrap").style.display="grid";document.getElementById("ganttWrap").style.display="none";};
document.getElementById("tabG").onclick=()=>{document.getElementById("tabG").classList.add("active");document.getElementById("tabT").classList.remove("active");document.getElementById("tasksWrap").style.display="none";document.getElementById("ganttWrap").style.display="block";renderGantt();};
["gWs","gRisk","gWin","gZoom","gMil","gIssue","gSlip"].forEach(id=>document.getElementById(id).addEventListener("change",renderGantt));
init();renderTasks();
<\/script></body></html>`;
}

main().catch(e => { console.error(e); process.exit(1); });
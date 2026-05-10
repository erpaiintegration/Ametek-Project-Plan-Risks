/* eslint-disable no-console */
const fs = require("node:fs");
const path = require("node:path");

const ROOT = process.cwd();
const STAGING_INPUT = path.join(ROOT, "imports", "staging", "scheduling-tower-data.json");
const DOCS_DIR = path.join(ROOT, "docs");
const OUT_HTML = path.join(DOCS_DIR, "scheduling-tower.html");
const OUT_JSON = path.join(DOCS_DIR, "scheduling-tower-data.json");

function esc(value) {
  return String(value ?? "")
    .replace(/&/g, "&amp;")
    .replace(/</g, "&lt;")
    .replace(/>/g, "&gt;")
    .replace(/\"/g, "&quot;")
    .replace(/'/g, "&#39;");
}

function renderHtml(data) {
  return `<!doctype html>
<html lang="en">
<head>
  <meta charset="UTF-8" />
  <meta name="viewport" content="width=device-width, initial-scale=1.0" />
  <title>Ametek Scheduling Tower</title>
  <style>
    :root{--bg:#f5f7fb;--card:#fff;--line:#dfe6f0;--txt:#0f172a;--muted:#475569;--blue:#2563eb;--amber:#d97706;--red:#dc2626;--green:#059669}
    body{margin:0;background:var(--bg);color:var(--txt);font:14px/1.45 Segoe UI,system-ui,-apple-system,sans-serif}
    header{padding:14px 18px;background:#0b1020;color:#f8fafc;display:flex;flex-wrap:wrap;gap:10px;align-items:center;justify-content:space-between}
    h1{font-size:18px;margin:0}
    .meta{font-size:12px;opacity:.9}
    main{padding:14px 18px;display:grid;gap:12px}
    .kpis{display:grid;grid-template-columns:repeat(auto-fill,minmax(180px,1fr));gap:10px}
    .kpi{background:var(--card);border:1px solid var(--line);border-radius:10px;padding:10px 12px}
    .kpi .label{font-size:11px;color:var(--muted);text-transform:uppercase;letter-spacing:.04em}
    .kpi .value{font-size:22px;font-weight:700}
    .kpi .value.red{color:var(--red)} .kpi .value.amber{color:var(--amber)} .kpi .value.green{color:var(--green)}
    .grid{display:grid;grid-template-columns:1fr 1fr;gap:12px}
    @media (max-width:1100px){.grid{grid-template-columns:1fr}}
    .panel{background:var(--card);border:1px solid var(--line);border-radius:10px;overflow:hidden}
    .panel h2{margin:0;padding:10px 12px;border-bottom:1px solid var(--line);font-size:13px}
    table{width:100%;border-collapse:collapse}
    th,td{padding:8px 10px;border-bottom:1px solid #edf2f7;text-align:left;vertical-align:top}
    th{font-size:11px;text-transform:uppercase;letter-spacing:.03em;color:var(--muted);background:#f8fafc}
    tr:nth-child(even) td{background:#fcfdff}
    .tag{padding:2px 6px;border-radius:999px;font-size:11px;border:1px solid #dbe2ea;background:#f8fafc;color:#334155}
    .small{font-size:12px;color:var(--muted)}
  </style>
</head>
<body>
  <header>
    <h1>📊 Ametek SAP S4 — Standalone Scheduling Tower</h1>
    <div class="meta" id="metaLine"></div>
  </header>
  <main>
    <section class="kpis" id="kpis"></section>
    <section class="grid">
      <article class="panel">
        <h2>Top slipped open tasks</h2>
        <div id="slippedTable"></div>
      </article>
      <article class="panel">
        <h2>Tasks due in next 14 days</h2>
        <div id="due14Table"></div>
      </article>
    </section>
    <article class="panel">
      <h2>Task list (first 400 rows)</h2>
      <div id="allTasksTable"></div>
    </article>
  </main>
<script>
async function loadData(){
  const res = await fetch('scheduling-tower-data.json',{cache:'no-store'});
  if(!res.ok) throw new Error('Failed to load scheduling-tower-data.json');
  return res.json();
}

function fmtDate(v){ if(!v) return '—'; const d=new Date(v); return Number.isNaN(d.getTime()) ? v : d.toLocaleDateString(); }
function num(v){ return Number(v||0).toLocaleString(); }

function buildKpis(m){
  const cards = [
    ['Total Tasks', num(m.totalTasks), ''],
    ['Open Tasks', num(m.openTasks), ''],
    ['Done Tasks', num(m.doneTasks), 'green'],
    ['Slipped Open', num(m.slippedOpen), m.slippedOpen>0?'red':'green'],
    ['Slip Rate', m.slipRatePct + '%', m.slipRatePct>=50?'red':m.slipRatePct>=25?'amber':'green'],
    ['Due 14d', num(m.due14), m.due14>0?'amber':'green'],
    ['Overdue Starts', num(m.overdueStarts), m.overdueStarts>0?'red':'green'],
    ['Open Critical', num(m.criticalTasks), m.criticalTasks>0?'amber':'green'],
    ['Open Milestones', num(m.milestonesOpen), m.milestonesOpen>0?'amber':'green']
  ];
  document.getElementById('kpis').innerHTML = cards.map(([label,val,cls])=>
    '<div class="kpi"><div class="label">'+label+'</div><div class="value '+cls+'">'+val+'</div></div>'
  ).join('');
}

function tableHtml(rows, cols){
  if(!rows.length) return '<div class="small" style="padding:10px 12px">No rows.</div>';
  const head = '<tr>'+cols.map(c=>'<th>'+c.label+'</th>').join('')+'</tr>';
  const body = rows.map(r=>'<tr>'+cols.map(c=>'<td>'+c.render(r)+'</td>').join('')+'</tr>').join('');
  return '<table><thead>'+head+'</thead><tbody>'+body+'</tbody></table>';
}

function render(data){
  const m = data.metrics || {};
  buildKpis(m);

  document.getElementById('metaLine').textContent =
    'Source: ' + (data.source?.projectName || 'MS Project') + ' · File: ' + (data.source?.file || 'n/a') +
    ' · Generated: ' + new Date(data.generatedAt).toLocaleString();

  const slippedCols = [
    {label:'UID', render:r=>String(r.uid)},
    {label:'Task', render:r=>r.name},
    {label:'Workstream / Resource', render:r=>r.resourceNames ? '<span class="tag">'+r.resourceNames+'</span>' : '—'},
    {label:'Finish', render:r=>fmtDate(r.finish)},
    {label:'Baseline', render:r=>fmtDate(r.baselineFinish)},
    {label:'Slip (d)', render:r=>String(r.slipDays ?? '—')}
  ];
  document.getElementById('slippedTable').innerHTML = tableHtml((data.lists?.topSlipped||[]).slice(0,40), slippedCols);

  const dueCols = [
    {label:'UID', render:r=>String(r.uid)},
    {label:'Task', render:r=>r.name},
    {label:'Finish', render:r=>fmtDate(r.finish)},
    {label:'% Complete', render:r=>String(r.percentComplete ?? 0) + '%'},
    {label:'Pred', render:r=>String(r.predCount||0)},
    {label:'Succ', render:r=>String(r.sucCount||0)}
  ];
  document.getElementById('due14Table').innerHTML = tableHtml((data.lists?.topDue14||[]).slice(0,40), dueCols);

  const allCols = [
    {label:'UID', render:r=>String(r.uid)},
    {label:'Name', render:r=>r.name},
    {label:'Lvl', render:r=>String(r.outlineLevel ?? '—')},
    {label:'Start', render:r=>fmtDate(r.start)},
    {label:'Finish', render:r=>fmtDate(r.finish)},
    {label:'Baseline Finish', render:r=>fmtDate(r.baselineFinish)},
    {label:'% Complete', render:r=>String(r.percentComplete ?? 0) + '%'},
    {label:'Slip', render:r=>r.slipDays == null ? '—' : String(r.slipDays)}
  ];
  document.getElementById('allTasksTable').innerHTML = tableHtml((data.tasks||[]).slice(0,400), allCols);
}

loadData().then(render).catch((err)=>{
  document.body.innerHTML = '<main style="padding:20px"><h2>Scheduling Tower failed to load</h2><pre>'+String(err.message||err)+'</pre></main>';
});
</script>
</body>
</html>`;
}

function main() {
  if (!fs.existsSync(STAGING_INPUT)) {
    throw new Error(`Missing input file: ${STAGING_INPUT}. Run ingest first.`);
  }

  const data = JSON.parse(fs.readFileSync(STAGING_INPUT, "utf8"));
  fs.mkdirSync(DOCS_DIR, { recursive: true });

  const html = renderHtml(data);
  fs.writeFileSync(OUT_HTML, html, "utf8");
  fs.writeFileSync(OUT_JSON, JSON.stringify(data), "utf8");

  console.log("Standalone Scheduling Tower generated:");
  console.log(`  HTML: ${OUT_HTML}`);
  console.log(`  Data: ${OUT_JSON}`);
}

main();

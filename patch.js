// patch2.js — Right-side Action Items panel
const fs = require('fs');
const FILE = 'schedule_intelligence/build.js';
let src = fs.readFileSync(FILE, 'utf8');
const orig = src;

function patch(desc, anchor, before) {
  if (!src.includes(anchor)) { console.error('MISS: ' + desc); return; }
  src = src.replace(anchor, before + anchor);
  console.log('✓ ' + desc);
}
function swap(desc, anchor, replacement) {
  if (!src.includes(anchor)) { console.error('MISS: ' + desc); return; }
  src = src.replace(anchor, replacement);
  console.log('✓ ' + desc);
}

// ── 1. CSS ────────────────────────────────────────────────────────────────
patch('action panel CSS', '/* NETWORK CANVAS DRAWER */', `/* ACTION PANEL */
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
`);

// ── 2. HTML (before #tooltip) ──────────────────────────────────────────────
patch('action panel HTML', '<div id="tooltip"></div>', `<button id="apToggleBtn" onclick="toggleActionPanel()" title="Toggle action items panel">
  Issues &amp; Risks<span class="ap-count-dot" id="apToggleDot">0</span>
</button>
<div id="actionPanel">
  <div class="ap-header">
    <span class="ap-title">Issues &amp; Risks</span>
    <span class="ap-count" id="apHeaderCount">0</span>
    <button onclick="closeActionPanel()">&#x2715;</button>
  </div>
  <div class="ap-sub" id="apSub">Showing all open items with issues</div>
  <div class="ap-lanes">
    <span class="ap-lane-chip active" data-lane="ALL" onclick="setApLane('ALL')">All</span>
    <span class="ap-lane-chip lane-now" data-lane="actNow" onclick="setApLane('actNow')">Act Now</span>
    <span class="ap-lane-chip lane-risk" data-lane="atRisk" onclick="setApLane('atRisk')">At Risk</span>
    <span class="ap-lane-chip lane-watch" data-lane="watchlist" onclick="setApLane('watchlist')">Watchlist</span>
  </div>
  <div id="apList"></div>
</div>
`);

// ── 3. JS (before openNetworkCanvas) ──────────────────────────────────────
patch('action panel JS', '// ── Network Canvas ──', `// ── Action Panel ────────────────────────────────────────────────────────
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
      taskUidSet = new Set([selTask.id]);
    }
  } else {
    // Apply view filter: build set of qualifying task IDs
    var vf = document.getElementById('viewFilter') ? document.getElementById('viewFilter').value : 'ALL';
    if (vf === 'CRITICAL') {
      taskUidSet = new Set();
      DATA.tasks.forEach(function(t) { if (t.cpm && t.cpm.critical && t.linkedRisks && t.linkedRisks.length) taskUidSet.add(t.id); });
    } else if (vf === 'DRIVER') {
      taskUidSet = new Set();
      DATA.tasks.forEach(function(t) { if ((t.cpmDriverScore > 0 || (t.cpm && t.cpm.critical)) && t.linkedRisks && t.linkedRisks.length) taskUidSet.add(t.id); });
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
    var linkedTask = a.taskId ? taskById.get(a.taskId) : null;
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
// ── Network Canvas ──`);

// ── 4. Hook into openNetworkCanvas to also refresh action panel ───────────
swap('hook action panel into canvas click',
  'selectedCanvasMilestones = new Set([String(uid)]);\n  document.getElementById(\'networkCanvas\').classList.add(\'open\');',
  'selectedCanvasMilestones = new Set([String(uid)]);\n  actionPanelTaskUid = String(uid);\n  if (actionPanelOpen) renderActionPanel();\n  document.getElementById(\'networkCanvas\').classList.add(\'open\');');

// ── 5. Hook setTimelineMode to refresh action panel ───────────────────────
swap('hook action panel into setTimelineMode',
  'function setTimelineMode(mode) {',
  'function setTimelineMode(mode) { actionPanelTaskUid = null; setTimeout(function(){ if(actionPanelOpen) renderActionPanel(); }, 50);');

// ── Write ─────────────────────────────────────────────────────────────────
if (src === orig) { console.error('No changes made.'); process.exit(1); }
fs.writeFileSync(FILE, src, 'utf8');
console.log('\nAll patches applied. Running build...');
require('child_process').execSync('node schedule_intelligence/build.js', { stdio: 'inherit' });
console.log('\nDone. Reload the browser.');
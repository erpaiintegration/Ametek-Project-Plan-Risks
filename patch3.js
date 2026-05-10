/* patch3.js – fix action panel: populate actionItems + fix taskById/id refs */
const fs   = require('fs');
const path = require('path');
const { execSync } = require('child_process');

const buildPath = path.join(__dirname, 'schedule_intelligence', 'build.js');
let src = fs.readFileSync(buildPath, 'utf8');

function swap(find, replace, label) {
  if (!src.includes(find)) { console.error('MISSING anchor: ' + label); process.exit(1); }
  src = src.replace(find, replace);
  console.log('✓ ' + label);
}

// ── Fix 1: derive actionItems from risk tiers + add to payload ────────────
swap(
`  const payload = {
    generatedAt: new Date().toISOString(),
    source: { file: path.basename(taskCsv) },
    enrichment,
    metrics,
    mermaid,
    laneAnchors,
    tasks: tasks.sort((a, b) => (a.outlineNumber || '').localeCompare(b.outlineNumber || '') || a.uid - b.uid)
  };`,
`  const actionItems = [
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
  };`,
'add actionItems to payload'
);

// ── Fix 2: taskById → taskByUid in renderActionPanel ─────────────────────
swap(
  `    var linkedTask = a.taskId ? taskById.get(a.taskId) : null;`,
  `    var linkedTask = a.taskId ? taskByUid.get(a.taskId) : null;`,
  'taskById → taskByUid'
);

// ── Fix 3: selTask.id → selTask.uid ──────────────────────────────────────
swap(
  `      taskUidSet = new Set([selTask.id]);`,
  `      taskUidSet = new Set([selTask.uid]);`,
  'selTask.id → selTask.uid'
);

// ── Fix 4: CRITICAL view filter – use t.uid, drop t.linkedRisks check ────
swap(
  `      DATA.tasks.forEach(function(t) { if (t.cpm && t.cpm.critical && t.linkedRisks && t.linkedRisks.length) taskUidSet.add(t.id); });`,
  `      DATA.tasks.forEach(function(t) { if (t.cpm && t.cpm.critical) taskUidSet.add(t.uid); });`,
  'CRITICAL filter fix'
);

// ── Fix 5: DRIVER view filter – use t.uid, drop t.linkedRisks check ──────
swap(
  `      DATA.tasks.forEach(function(t) { if ((t.cpmDriverScore > 0 || (t.cpm && t.cpm.critical)) && t.linkedRisks && t.linkedRisks.length) taskUidSet.add(t.id); });`,
  `      DATA.tasks.forEach(function(t) { if (t.cpm && t.cpm.critical) taskUidSet.add(t.uid); });`,
  'DRIVER filter fix'
);

fs.writeFileSync(buildPath, src, 'utf8');
console.log('\npatch3.js applied — building...\n');
execSync('node schedule_intelligence/build.js', { stdio: 'inherit', cwd: __dirname });

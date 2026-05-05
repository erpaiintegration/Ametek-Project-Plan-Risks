/* eslint-disable no-console */
require("dotenv/config");

const { Client } = require("@notionhq/client");

const notion = new Client({ auth: process.env.NOTION_API_KEY });

const TASKS_DB_ID = process.env.NOTION_TASKS_DB_ID;
const RISKS_DB_ID = process.env.NOTION_RISKS_DB_ID || "357ae9be-8a60-8190-b720-c130c7104cf1";
const PROJECT_PARENT_PAGE_ID =
  process.env.NOTION_AMETEK_PROJECT_PAGE_ID || "62bae9be-8a60-8398-b35a-81cf9997daa0";
const SCHEDULE_DASHBOARD_PAGE_ID =
  process.env.NOTION_SCHEDULE_DASHBOARD_PAGE_ID || "357ae9be-8a60-81ce-805c-efb099469f41";

const TASKS_URL = `https://www.notion.so/${String(TASKS_DB_ID).replace(/-/g, "")}`;
const RISKS_URL = `https://www.notion.so/${String(RISKS_DB_ID).replace(/-/g, "")}`;

function text(parts) {
  return (parts || []).map((p) => p.plain_text).join("");
}

function title(prop) {
  return prop?.type === "title" ? text(prop.title || []) : "";
}

function rich(prop) {
  return prop?.type === "rich_text" ? text(prop.rich_text || []) : "";
}

function dateValue(prop) {
  return prop?.type === "date" ? prop.date?.start || null : null;
}

function numberValue(prop) {
  return prop?.type === "number" ? prop.number : null;
}

function checkboxValue(prop) {
  return prop?.type === "checkbox" ? !!prop.checkbox : false;
}

function relationCount(prop) {
  return prop?.type === "relation" ? (prop.relation || []).length : 0;
}

function selectValue(prop) {
  if (prop?.type === "select") return prop.select?.name || "";
  if (prop?.type === "status") return prop.status?.name || "";
  return "";
}

function pageTitleFromProperties(properties) {
  const props = properties || {};
  for (const key of Object.keys(props)) {
    const p = props[key];
    if (p?.type === "title") {
      const t = title(p).trim();
      if (t) return t;
    }
  }
  return "(Untitled)";
}

function relationCountByNames(properties, names) {
  for (const name of names) {
    const p = properties[name];
    if (p?.type === "relation") return relationCount(p);
  }
  return 0;
}

function daysDiff(current, baseline) {
  if (!current || !baseline) return null;
  const currentDate = new Date(current);
  const baselineDate = new Date(baseline);
  if (Number.isNaN(currentDate.getTime()) || Number.isNaN(baselineDate.getTime())) {
    return null;
  }
  return Math.round((currentDate - baselineDate) / (1000 * 60 * 60 * 24));
}

function normalizeKey(value) {
  return String(value || "")
    .trim()
    .toLowerCase()
    .replace(/\s+/g, " ");
}

function rt(content, options = {}) {
  const annotations = options.annotations || {};
  const textPayload = { content };
  if (options.url) {
    textPayload.link = { url: options.url };
  }
  return { type: "text", text: textPayload, annotations };
}

function buildParagraph(textContent) {
  return {
    object: "block",
    type: "paragraph",
    paragraph: { rich_text: [rt(textContent)] }
  };
}

function buildHeading2(textContent) {
  return {
    object: "block",
    type: "heading_2",
    heading_2: { rich_text: [rt(textContent)] }
  };
}

function buildCallout(emoji, color, textContent) {
  return {
    object: "block",
    type: "callout",
    callout: {
      icon: { type: "emoji", emoji },
      color,
      rich_text: [rt(textContent)]
    }
  };
}

async function fetchAllPages(databaseId, label) {
  const pages = [];
  let cursor;
  do {
    const response = await notion.databases.query({
      database_id: databaseId,
      start_cursor: cursor,
      page_size: 100
    });
    pages.push(...response.results.filter((r) => r.object === "page"));
    cursor = response.has_more ? response.next_cursor : undefined;
  } while (cursor);
  console.log(`${label}: ${pages.length} rows loaded`);
  return pages;
}

async function listChildren(blockId) {
  const blocks = [];
  let cursor;
  do {
    const response = await notion.blocks.children.list({
      block_id: blockId,
      start_cursor: cursor,
      page_size: 100
    });
    blocks.push(...response.results);
    cursor = response.has_more ? response.next_cursor : undefined;
  } while (cursor);
  return blocks;
}

async function replacePageContent(pageId, children) {
  const existing = await listChildren(pageId);
  for (const block of existing) {
    await notion.blocks.delete({ block_id: block.id });
  }

  const chunkSize = 80;
  for (let i = 0; i < children.length; i += chunkSize) {
    const chunk = children.slice(i, i + chunkSize);
    await notion.blocks.children.append({ block_id: pageId, children: chunk });
  }
}

async function findOrCreateChildPage(parentPageId, pageTitle, emoji) {
  const search = await notion.search({ query: pageTitle, page_size: 50 });
  const existing = search.results.find(
    (r) =>
      r.object === "page" &&
      r.parent?.type === "page_id" &&
      r.parent?.page_id === parentPageId &&
      r.properties?.title?.type === "title" &&
      text(r.properties.title.title || []) === pageTitle
  );

  if (existing) {
    return existing;
  }

  return notion.pages.create({
    parent: { type: "page_id", page_id: parentPageId },
    icon: { type: "emoji", emoji },
    properties: {
      title: {
        title: [{ type: "text", text: { content: pageTitle } }]
      }
    }
  });
}

async function findOrCreateSnapshotDb(parentPageId) {
  const dbName = "Ametek Schedule KPI Snapshots";
  const search = await notion.search({
    query: dbName,
    filter: { property: "object", value: "database" },
    page_size: 50
  });

  const existing = search.results.find(
    (r) =>
      r.object === "database" &&
      text(r.title || []) === dbName &&
      r.parent?.type === "page_id" &&
      r.parent?.page_id === parentPageId
  );

  if (existing) return existing;

  return notion.databases.create({
    parent: { type: "page_id", page_id: parentPageId },
    title: [{ type: "text", text: { content: dbName } }],
    icon: { type: "emoji", emoji: "🧾" },
    properties: {
      Date: { title: {} },
      "Total Tasks": { number: {} },
      "Open Tasks": { number: {} },
      "Done Tasks": { number: {} },
      "Slipped Open": { number: {} },
      "Slip Rate %": { number: {} },
      "Due 14d": { number: {} },
      "Overdue Starts": { number: {} },
      Health: {
        select: {
          options: [{ name: "Red" }, { name: "Amber" }, { name: "Green" }]
        }
      }
    }
  });
}

async function upsertTodaySnapshot(snapshotDbId, metrics) {
  const today = new Date().toISOString().slice(0, 10);
  const query = await notion.databases.query({
    database_id: snapshotDbId,
    filter: {
      property: "Date",
      title: { equals: today }
    },
    page_size: 10
  });

  const properties = {
    Date: { title: [{ type: "text", text: { content: today } }] },
    "Total Tasks": { number: metrics.total },
    "Open Tasks": { number: metrics.open },
    "Done Tasks": { number: metrics.done },
    "Slipped Open": { number: metrics.slippedOpen },
    "Slip Rate %": { number: Number(metrics.slipRateOpen.toFixed(1)) },
    "Due 14d": { number: metrics.due14 },
    "Overdue Starts": { number: metrics.overdueStarts },
    Health: { select: { name: metrics.healthLabel } }
  };

  if (query.results.length > 0) {
    await notion.pages.update({ page_id: query.results[0].id, properties });
  } else {
    await notion.pages.create({
      parent: { database_id: snapshotDbId },
      properties
    });
  }
}

async function getLatestSnapshot(snapshotDbId, { excludeDate } = {}) {
  const response = await notion.databases.query({
    database_id: snapshotDbId,
    sorts: [{ property: "Date", direction: "descending" }],
    page_size: 10
  });

  const rows = response.results.filter((r) => {
    const dateText = title((r.properties || {}).Date);
    return !excludeDate || dateText !== excludeDate;
  });

  if (!rows.length) return null;

  const props = rows[0].properties || {};
  return {
    date: title(props.Date),
    total: numberValue(props["Total Tasks"]) ?? 0,
    open: numberValue(props["Open Tasks"]) ?? 0,
    done: numberValue(props["Done Tasks"]) ?? 0,
    slippedOpen: numberValue(props["Slipped Open"]) ?? 0,
    slipRateOpen: numberValue(props["Slip Rate %"]) ?? 0,
    due14: numberValue(props["Due 14d"]) ?? 0,
    overdueStarts: numberValue(props["Overdue Starts"]) ?? 0,
    healthLabel:
      props.Health?.select?.name ||
      props.Health?.status?.name ||
      "Unknown"
  };
}

function delta(current, previous) {
  if (previous == null || Number.isNaN(previous)) return null;
  return current - previous;
}

function fmtSigned(value, suffix = "") {
  if (value == null) return "n/a";
  const sign = value > 0 ? "+" : "";
  return `${sign}${value}${suffix}`;
}

function trendWord(value, improveDirection = "down") {
  if (value == null || value === 0) return "flat";
  if (improveDirection === "down") return value < 0 ? "improving" : "worsening";
  return value > 0 ? "improving" : "worsening";
}

function buildKpiCard(label, value, deltaValue, options = {}) {
  const suffix = options.suffix || "";
  const improveDirection = options.improveDirection || "down";
  const deltaText =
    deltaValue == null
      ? "vs previous snapshot: n/a"
      : `vs previous snapshot: ${fmtSigned(Number(deltaValue.toFixed?.(1) ?? deltaValue), suffix)}`;
  const trend = trendWord(deltaValue, improveDirection);

  return buildCallout(
    "🧩",
    options.color || "gray_background",
    `${label}: ${value}${suffix}\n${deltaText}\nTrend: ${trend}`
  );
}

function buildSoWhatSummary(metrics, topWorkstreams) {
  const top3 = topWorkstreams.slice(0, 3).map((x) => `${x.workstream} (${x.count})`).join(", ");
  const pressureText = top3 || "No concentrated workstream pressure detected yet";

  const statusLine =
    metrics.healthLabel === "Red"
      ? "The schedule is in intervention mode. Current slippage is beyond agreed tolerance."
      : metrics.healthLabel === "Amber"
      ? "The schedule is at risk and needs focused recovery actions this week."
      : "The schedule is within tolerance, but watch near-term workload and risk signals.";

  const meaning = [
    `1) Overall status: ${metrics.healthEmoji} ${metrics.healthLabel}. ${statusLine}`,
    `2) Main problem: ${metrics.slippedOpen} of ${metrics.open} open tasks are late vs baseline (${metrics.slipRateOpen.toFixed(
      1
    )}% slip rate).`,
    `3) Near-term risk: ${metrics.due14} tasks are due in the next 14 days; this is the likely next wave of misses if not triaged.`,
    `4) Execution friction: ${metrics.overdueStarts} tasks have not started even though their planned start date passed.`,
    `5) Concentrated pressure: ${pressureText}.`
  ];

  const actions = [
    "This week actions:",
    "• Freeze non-critical scope in top pressure workstreams.",
    "• Daily review of due-in-14 and overdue-start lists with named owners.",
    "• Approve date re-baselines only with dependency impact visible.",
    "• Escalate blockers that have no owner or due date within 48 hours."
  ];

  const definitions = [
    "Metric definitions:",
    "• Slipped Open = task is not complete and Finish date is later than Baseline Finish.",
    "• Slip Rate = Slipped Open / Open Tasks.",
    "• Due 14d = incomplete tasks with Finish date in next 14 days.",
    "• Overdue Starts = 0% complete tasks whose Start date is in the past."
  ];

  return [...meaning, "", ...actions, "", ...definitions].join("\n");
}

function buildDependencyImpactBasics(metrics) {
  const dependencyRatio = metrics.dependencyChangeTasks
    ? (metrics.potentialImpactedSuccessors / metrics.dependencyChangeTasks).toFixed(2)
    : "0.00";

  const milestoneSignal =
    metrics.slippedMilestones > 0
      ? `Milestone impact is active: ${metrics.slippedMilestones} slipped milestone(s) currently affect ${metrics.milestoneImpactedSuccessors} downstream successor task links.`
      : "Milestone impact is currently low in this snapshot: no slipped milestones with successor links were detected.";

  return [
    "Basic dependency explanation:",
    "• Upstream = predecessor tasks (work that must finish first).",
    "• Downstream = successor tasks (work that depends on predecessors).",
    "",
    "What the issue is:",
    `• ${metrics.dependencyChangeTasks} slipped tasks are dependency-linked, meaning late tasks are not isolated—they sit inside dependency chains.`,
    `• These late tasks expose ${metrics.potentialImpactedSuccessors} downstream successor links (${dependencyRatio} successor links per slipped dependency-linked task on average).`,
    "",
    "What up/down impact means operationally:",
    "• If an upstream task slips, downstream tasks either start late, compress, or miss dates.",
    "• The larger the downstream successor count, the larger the schedule blast radius.",
    "",
    "What this means to milestones:",
    `• ${milestoneSignal}`,
    "• A milestone slip means phase-completion confidence drops and cutover/readiness windows become higher risk.",
    "• Even when a milestone itself is not yet slipped, heavy upstream/downstream pressure is an early warning that milestones can move next."
  ].join("\n");
}

function dedupeTop(items, keyFn, limit = 10) {
  const seen = new Set();
  const out = [];
  for (const item of items) {
    const key = keyFn(item);
    if (seen.has(key)) continue;
    seen.add(key);
    out.push(item);
    if (out.length >= limit) break;
  }
  return out;
}

function computeMetrics(tasks, risks) {
  const now = new Date();
  const in14 = new Date(now.getTime() + 14 * 24 * 60 * 60 * 1000);
  const in30 = new Date(now.getTime() + 30 * 24 * 60 * 60 * 1000);

  const metrics = {
    total: tasks.length,
    open: 0,
    done: 0,
    slippedOpen: 0,
    overdueStarts: 0,
    due14: 0,
    milestones30: 0,
    tasksWithRisk: 0,
    slipDaysTotal: 0,
    maxSlip: 0,
    openIssues: 0,
    openRisks: 0
  };

  const slipped = [];
  const overdueStart = [];
  const dueSoon = [];
  const milestonesSoon = [];
  const slippedMilestones = [];
  const dependencyChanges = [];
  const workstreamPressure = new Map();
  const openIssues = [];
  const riskTypeCounts = new Map();

  for (const task of tasks) {
    const p = task.properties || {};
    const uid = p["Unique ID"]?.number ?? null;
    const name = title(p["Task Name"]);
    const workstream = rich(p.Workstream).trim() || "(Unassigned)";
    const start = dateValue(p.Start);
    const finish = dateValue(p.Finish);
    const baselineFinish = dateValue(p["Baseline Finish"]) || dateValue(p["Baseline10 Finish"]);
    const percentComplete = numberValue(p["% Complete"]) ?? 0;
    const milestone = checkboxValue(p.Milestone);
    const predecessorCount = relationCountByNames(p, ["Predecessor Tasks", "Predecessors"]);
    const successorCount = relationCountByNames(p, ["Successor Tasks", "Successors"]);

    if (percentComplete >= 100) metrics.done += 1;
    else metrics.open += 1;

    const slipDays = daysDiff(finish, baselineFinish);
    if (slipDays != null && slipDays > 0 && percentComplete < 100) {
      metrics.slippedOpen += 1;
      metrics.slipDaysTotal += slipDays;
      metrics.maxSlip = Math.max(metrics.maxSlip, slipDays);
      slipped.push({ uid, name, workstream, finish, start, percentComplete, slipDays });
      workstreamPressure.set(workstream, (workstreamPressure.get(workstream) || 0) + 1);

      if (milestone) {
        slippedMilestones.push({
          uid,
          name,
          workstream,
          finish,
          baselineFinish,
          slipDays,
          predecessorCount,
          successorCount
        });
      }

      if (predecessorCount > 0 || successorCount > 0 || milestone) {
        dependencyChanges.push({
          uid,
          name,
          workstream,
          finish,
          baselineFinish,
          slipDays,
          predecessorCount,
          successorCount,
          milestone,
          impactScore: successorCount * Math.max(1, slipDays)
        });
      }
    }

    if (percentComplete === 0 && start) {
      const s = new Date(start);
      if (!Number.isNaN(s.getTime()) && s < now) {
        metrics.overdueStarts += 1;
        overdueStart.push({ uid, name, workstream, start, finish, percentComplete });
      }
    }

    if (percentComplete < 100 && finish) {
      const f = new Date(finish);
      if (!Number.isNaN(f.getTime()) && f >= now && f <= in14) {
        metrics.due14 += 1;
        dueSoon.push({ uid, name, workstream, finish, slipDays: slipDays || 0, percentComplete });
      }
      if (milestone && !Number.isNaN(f.getTime()) && f >= now && f <= in30) {
        metrics.milestones30 += 1;
        milestonesSoon.push({ uid, name, workstream, finish, slipDays: slipDays || 0, percentComplete });
      }
    }

    if ((p["Risks & Issues"]?.relation || []).length > 0) {
      metrics.tasksWithRisk += 1;
    }
  }

  for (const risk of risks) {
    const p = risk.properties || {};
    const status = p.Status?.status?.name || p.Status?.select?.name || "";
    if (/resolved|closed/i.test(status)) continue;
    const type = p.Type?.select?.name || "Risk";
    const category =
      selectValue(p["Risk Type"]) ||
      selectValue(p.Category) ||
      selectValue(p.Family) ||
      "(Uncategorized)";
    const severity =
      selectValue(p.Severity) ||
      selectValue(p.Priority) ||
      selectValue(p.Impact) ||
      "(Unrated)";
    const impactedTasks = relationCountByNames(p, [
      "Tasks",
      "Related Tasks",
      "Impacted Tasks",
      "Task",
      "Task Links"
    ]);

    if (/issue/i.test(type)) metrics.openIssues += 1;
    else metrics.openRisks += 1;

    const typeKey = `${type} — ${category}`;
    riskTypeCounts.set(typeKey, (riskTypeCounts.get(typeKey) || 0) + 1);

    if (/issue/i.test(type)) {
      const severityRank =
        /critical/i.test(severity) ? 4 : /high/i.test(severity) ? 3 : /medium/i.test(severity) ? 2 : 1;
      openIssues.push({
        name: pageTitleFromProperties(p),
        severity,
        category,
        status,
        impactedTasks,
        rankScore: severityRank * 1000 + impactedTasks
      });
    }
  }

  metrics.progressPct = metrics.total ? (metrics.done / metrics.total) * 100 : 0;
  metrics.slipRateOpen = metrics.open ? (metrics.slippedOpen / metrics.open) * 100 : 0;
  metrics.avgSlipDays = metrics.slippedOpen ? metrics.slipDaysTotal / metrics.slippedOpen : 0;

  const rag = {
    red:
      metrics.slippedOpen >= 1200 ||
      metrics.slipRateOpen >= 75 ||
      metrics.due14 >= 100 ||
      metrics.overdueStarts >= 100,
    amber:
      metrics.slippedOpen >= 700 ||
      metrics.slipRateOpen >= 50 ||
      metrics.due14 >= 60 ||
      metrics.overdueStarts >= 40
  };

  if (rag.red) {
    metrics.healthEmoji = "🔴";
    metrics.healthLabel = "Red";
  } else if (rag.amber) {
    metrics.healthEmoji = "🟠";
    metrics.healthLabel = "Amber";
  } else {
    metrics.healthEmoji = "🟢";
    metrics.healthLabel = "Green";
  }

  slipped.sort((a, b) => b.slipDays - a.slipDays || (a.finish || "").localeCompare(b.finish || ""));
  overdueStart.sort((a, b) => (a.start || "").localeCompare(b.start || ""));
  dueSoon.sort((a, b) => (a.finish || "").localeCompare(b.finish || "") || b.slipDays - a.slipDays);
  milestonesSoon.sort((a, b) => (a.finish || "").localeCompare(b.finish || ""));
  dependencyChanges.sort((a, b) => b.impactScore - a.impactScore || b.slipDays - a.slipDays);
  slippedMilestones.sort((a, b) => b.successorCount - a.successorCount || b.slipDays - a.slipDays);
  openIssues.sort((a, b) => b.rankScore - a.rankScore || b.impactedTasks - a.impactedTasks);

  const topWorkstreams = [...workstreamPressure.entries()]
    .sort((a, b) => b[1] - a[1])
    .slice(0, 10)
    .map(([workstream, count]) => ({ workstream, count }));

  const topSlipped = dedupeTop(
    slipped,
    (x) => `${normalizeKey(x.name)}|${normalizeKey(x.workstream)}`,
    10
  );

  const topOverdueStarts = dedupeTop(
    overdueStart,
    (x) => `${normalizeKey(x.name)}|${normalizeKey(x.workstream)}`,
    10
  );

  const topDue14 = dedupeTop(
    dueSoon,
    (x) => `${normalizeKey(x.name)}|${normalizeKey(x.workstream)}`,
    10
  );

  const topMilestones30 = dedupeTop(
    milestonesSoon,
    (x) => `${normalizeKey(x.name)}|${normalizeKey(x.workstream)}`,
    10
  );

  const topIssues = dedupeTop(openIssues, (x) => normalizeKey(x.name), 10);

  const riskTypeBreakdown = [...riskTypeCounts.entries()]
    .sort((a, b) => b[1] - a[1])
    .map(([typeKey, count]) => ({ typeKey, count }));

  const topDependencyChanges = dedupeTop(
    dependencyChanges,
    (x) => `${x.uid}|${normalizeKey(x.name)}`,
    15
  );

  const topMilestoneImpacts = dedupeTop(
    slippedMilestones,
    (x) => `${x.uid}|${normalizeKey(x.name)}`,
    10
  );

  metrics.slippedMilestones = slippedMilestones.length;
  metrics.dependencyChangeTasks = dependencyChanges.length;
  metrics.potentialImpactedSuccessors = dependencyChanges.reduce((sum, x) => sum + x.successorCount, 0);
  metrics.milestoneImpactedSuccessors = slippedMilestones.reduce((sum, x) => sum + x.successorCount, 0);

  const decisions = [];
  if (topWorkstreams.length > 0) {
    const top3 = topWorkstreams.slice(0, 3).map((w) => w.workstream).join(", ");
    decisions.push(
      `Approve a focused schedule recovery plan for ${top3} with weekly accountable owners and target date moves.`
    );
  }
  if (metrics.due14 >= 60) {
    decisions.push(
      `Prioritize resource allocation for the next 14-day window (${metrics.due14} tasks due) and freeze non-critical work.`
    );
  }
  if (metrics.openIssues + metrics.openRisks >= 10) {
    decisions.push(
      `Escalate RAID governance: enforce owner + due date on all open records and review closure progress in weekly steering.`
    );
  }
  while (decisions.length < 3) {
    decisions.push(
      "Confirm baseline integrity and critical-path tagging in MSP export to improve confidence in schedule decision signals."
    );
  }

  return {
    metrics,
    topIssues,
    riskTypeBreakdown,
    topSlipped,
    topOverdueStarts,
    topDue14,
    topMilestones30,
    topDependencyChanges,
    topMilestoneImpacts,
    topWorkstreams,
    decisions: decisions.slice(0, 3)
  };
}

function listText(items, mapFn, emptyText) {
  if (!items.length) return emptyText;
  return items.map((item, i) => `${i + 1}. ${mapFn(item)}`).join("\n");
}

async function main() {
  if (!process.env.NOTION_API_KEY || !TASKS_DB_ID) {
    throw new Error("NOTION_API_KEY and NOTION_TASKS_DB_ID are required.");
  }

  const [tasks, risks] = await Promise.all([
    fetchAllPages(TASKS_DB_ID, "tasks"),
    fetchAllPages(RISKS_DB_ID, "risks")
  ]);

  const data = computeMetrics(tasks, risks);
  const m = data.metrics;

  const today = new Date().toISOString().slice(0, 10);
  const snapshotDb = await findOrCreateSnapshotDb(PROJECT_PARENT_PAGE_ID);
  const previousSnapshot = await getLatestSnapshot(snapshotDb.id, { excludeDate: today });

  const deltas = {
    open: delta(m.open, previousSnapshot?.open),
    slippedOpen: delta(m.slippedOpen, previousSnapshot?.slippedOpen),
    slipRateOpen: delta(Number(m.slipRateOpen.toFixed(1)), previousSnapshot?.slipRateOpen),
    due14: delta(m.due14, previousSnapshot?.due14),
    overdueStarts: delta(m.overdueStarts, previousSnapshot?.overdueStarts),
    tasksWithRisk: delta(m.tasksWithRisk, previousSnapshot?.tasksWithRisk)
  };

  const summarySentence = previousSnapshot
    ? `Schedule is ${m.healthLabel} (${m.healthEmoji}). Slip rate is ${m.slipRateOpen.toFixed(1)}% (${fmtSigned(
      Number((deltas.slipRateOpen ?? 0).toFixed(1)),
      "%"
    )} vs ${previousSnapshot.date}), with ${m.due14} tasks due in 14 days (${fmtSigned(deltas.due14)}) and ${m.overdueStarts} overdue starts (${fmtSigned(deltas.overdueStarts)}).`
    : `Schedule is ${m.healthLabel} (${m.healthEmoji}). Baseline snapshot established with ${m.slippedOpen} slipped open tasks, ${m.due14} due in 14 days, and ${m.overdueStarts} overdue starts.`;

  const slippedText = listText(
    data.topSlipped,
    (x) => `${x.uid} — ${x.name} | ${x.workstream} | ${x.slipDays}d late`,
    "No slipped open tasks."
  );
  const overdueText = listText(
    data.topOverdueStarts,
    (x) => `${x.uid} — ${x.name} | ${x.workstream} | start ${x.start}`,
    "No overdue unstarted tasks."
  );
  const due14Text = listText(
    data.topDue14,
    (x) => `${x.uid} — ${x.name} | ${x.workstream} | due ${x.finish}`,
    "No tasks due in the next 14 days."
  );
  const milestoneText = listText(
    data.topMilestones30,
    (x) => `${x.uid} — ${x.name} | ${x.workstream} | ${x.finish}`,
    "No incomplete milestones due in next 30 days."
  );
  const topIssuesText = listText(
    data.topIssues,
    (x) => `${x.name} | severity: ${x.severity} | type: ${x.category} | impacted tasks: ${x.impactedTasks}`,
    "No open issues."
  );
  const riskTypeText = listText(
    data.riskTypeBreakdown,
    (x) => `${x.typeKey} — ${x.count}`,
    "No risk-type records found."
  );
  const dependencySummaryText = [
    `Dependency-change tasks (slipped with predecessor/successor links): ${m.dependencyChangeTasks}`,
    `Potential downstream impacted successors: ${m.potentialImpactedSuccessors}`,
    `Slipped milestones: ${m.slippedMilestones}`,
    `Milestone successor impacts: ${m.milestoneImpactedSuccessors}`
  ].join("\n");
  const dependencyChangeText = listText(
    data.topDependencyChanges,
    (x) => `${x.uid} — ${x.name} | ${x.slipDays}d late | pred ${x.predecessorCount} | succ ${x.successorCount} | impact ${x.impactScore}`,
    "No dependency-linked slipped tasks found."
  );
  const milestoneImpactText = listText(
    data.topMilestoneImpacts,
    (x) => `${x.uid} — ${x.name} | ${x.slipDays}d late | baseline ${x.baselineFinish || "n/a"} -> finish ${x.finish || "n/a"} | succ impacted ${x.successorCount}`,
    "No slipped milestones with dependency impacts found."
  );
  const wsText = listText(
    data.topWorkstreams,
    (x) => `${x.workstream} — ${x.count} slipped open tasks`,
    "No workstream slippage."
  );
  const decisionsText = data.decisions.map((d, i) => `${i + 1}. ${d}`).join("\n");
  const soWhatText = buildSoWhatSummary(m, data.topWorkstreams);
  const dependencyBasicsText = buildDependencyImpactBasics(m);

  const dashboardBlocks = [
    {
      object: "block",
      type: "heading_1",
      heading_1: { rich_text: [rt("Ametek SAP S4 — Task Scheduling Dashboard")] }
    },
    {
      object: "block",
      type: "paragraph",
      paragraph: {
        rich_text: [
          rt("Traditional PM schedule dashboard (tasks only). Updated: "),
          rt(today, { annotations: { bold: true } })
        ]
      }
    },
    buildHeading2("What this tells us right now"),
    buildCallout("🧠", "yellow_background", soWhatText),
    buildCallout("🗣️", "gray_background", summarySentence),
    { object: "block", type: "divider", divider: {} },
    buildHeading2("KPI cards (snapshot + delta)"),
    buildKpiCard("Open Tasks", m.open, deltas.open, { color: "gray_background", improveDirection: "down" }),
    buildKpiCard("Slipped Open", m.slippedOpen, deltas.slippedOpen, {
      color: "orange_background",
      improveDirection: "down"
    }),
    buildKpiCard("Slip Rate", Number(m.slipRateOpen.toFixed(1)), deltas.slipRateOpen, {
      color: "orange_background",
      suffix: "%",
      improveDirection: "down"
    }),
    buildKpiCard("Due in 14 Days", m.due14, deltas.due14, {
      color: "blue_background",
      improveDirection: "down"
    }),
    buildKpiCard("Overdue Starts", m.overdueStarts, deltas.overdueStarts, {
      color: "red_background",
      improveDirection: "down"
    }),
    buildKpiCard("Tasks with Risk/Issue links", m.tasksWithRisk, null, {
      color: "purple_background",
      improveDirection: "up"
    }),
    buildKpiCard("Open Issues", m.openIssues, null, {
      color: "red_background",
      improveDirection: "down"
    }),
    {
      object: "block",
      type: "callout",
      callout: {
        icon: { type: "emoji", emoji: "🧭" },
        color: "yellow_background",
        rich_text: [
          rt(`Schedule health: ${m.healthEmoji} ${m.healthLabel}\n`, { annotations: { bold: true } }),
          rt(
            `Thresholds: Red if slipped open >= 1200 OR slip rate >= 75% OR due 14d >= 100 OR overdue starts >= 100; Amber if slipped open >= 700 OR slip rate >= 50% OR due 14d >= 60 OR overdue starts >= 40.\n`
          ),
          rt(`Total: ${m.total} | Open: ${m.open} | Done: ${m.done}\n`),
          rt(
            `Slipped open: ${m.slippedOpen} | Slip rate on open: ${m.slipRateOpen.toFixed(1)}% | Avg slip: ${m.avgSlipDays.toFixed(1)}d | Max slip: ${m.maxSlip}d\n`
          ),
          rt(`Overdue starts: ${m.overdueStarts} | Due next 14d: ${m.due14} | Milestones next 30d: ${m.milestones30}`)
        ]
      }
    },
    buildHeading2("Trend & change notes"),
    buildCallout(
      "📈",
      "gray_background",
      previousSnapshot
        ? [
            `Compared to ${previousSnapshot.date}:`,
            `• Open tasks ${fmtSigned(deltas.open)}`,
            `• Slipped open ${fmtSigned(deltas.slippedOpen)}`,
            `• Slip rate ${fmtSigned(Number((deltas.slipRateOpen ?? 0).toFixed(1)), "%")}`,
            `• Due in 14 days ${fmtSigned(deltas.due14)}`,
            `• Overdue starts ${fmtSigned(deltas.overdueStarts)}`
          ].join("\n")
        : "No prior snapshot exists yet for trend comparison."
    ),
      buildHeading2("Top issues (open)"),
      buildCallout("🚨", "red_background", topIssuesText),
      buildHeading2("Risk type breakdown"),
      buildCallout("🧷", "gray_background", riskTypeText),
      buildHeading2("Dependency change and milestone impact"),
      buildCallout("🔗", "yellow_background", dependencySummaryText),
      buildHeading2("How upstream/downstream impact works"),
      buildCallout("🧭", "yellow_background", dependencyBasicsText),
      buildHeading2("Top dependency changes (slip × successors)"),
      buildCallout("🔄", "orange_background", dependencyChangeText),
      buildHeading2("Milestone dependency impacts"),
      buildCallout("🏁", "blue_background", milestoneImpactText),
    buildHeading2("Top 10 slipped open tasks (deduped)"),
    buildCallout("⚠️", "orange_background", slippedText),
    buildHeading2("Top 10 overdue starts (deduped)"),
    buildCallout("⏳", "red_background", overdueText),
    buildHeading2("Top 10 tasks due in next 14 days (deduped)"),
    buildCallout("📌", "blue_background", due14Text),
    buildHeading2("Top 10 upcoming milestones in next 30 days (deduped)"),
    buildCallout("🏁", "blue_background", milestoneText),
    buildHeading2("Top 10 workstreams by schedule pressure"),
    buildCallout("📊", "gray_background", wsText),
    buildHeading2("Top 3 decisions needed this week"),
    buildCallout("🧠", "purple_background", decisionsText),
    { object: "block", type: "divider", divider: {} },
    buildParagraph("Drill-down links:"),
    {
      object: "block",
      type: "paragraph",
      paragraph: {
        rich_text: [
          rt("Tasks database", { url: TASKS_URL, annotations: { bold: true } }),
          rt(" • "),
          rt("Risks & Issues", { url: RISKS_URL, annotations: { bold: true } })
        ]
      }
    }
  ];

  await replacePageContent(SCHEDULE_DASHBOARD_PAGE_ID, dashboardBlocks);

  await upsertTodaySnapshot(snapshotDb.id, m);

  const steeringPage = await findOrCreateChildPage(
    PROJECT_PARENT_PAGE_ID,
    "Ametek SAP S4 — Steering Committee Schedule View",
    "🧑‍⚖️"
  );
  const pmoPage = await findOrCreateChildPage(
    PROJECT_PARENT_PAGE_ID,
    "Ametek SAP S4 — PMO Schedule Working View",
    "🛠️"
  );

  await replacePageContent(steeringPage.id, [
    {
      object: "block",
      type: "heading_1",
      heading_1: { rich_text: [rt("Steering Committee — Schedule View")] }
    },
    buildParagraph(`Updated ${today}`),
    buildHeading2("What this tells us right now"),
    buildCallout("🧠", "yellow_background", soWhatText),
    buildCallout(
      "🧭",
      "yellow_background",
      `Overall schedule health: ${m.healthEmoji} ${m.healthLabel}. Open ${m.open}/${m.total}; slipped open ${m.slippedOpen}; due next 14d ${m.due14}.`
    ),
    buildCallout("🗣️", "gray_background", summarySentence),
    buildHeading2("KPI movement vs previous snapshot"),
    buildCallout(
      "📈",
      "gray_background",
      previousSnapshot
        ? [
            `Reference date: ${previousSnapshot.date}`,
            `• Slipped open ${fmtSigned(deltas.slippedOpen)}`,
            `• Slip rate ${fmtSigned(Number((deltas.slipRateOpen ?? 0).toFixed(1)), "%")}`,
            `• Due in 14 days ${fmtSigned(deltas.due14)}`,
            `• Overdue starts ${fmtSigned(deltas.overdueStarts)}`
          ].join("\n")
        : "No prior snapshot exists yet for trend comparison."
    ),
    buildHeading2("Decisions needed this week"),
    buildCallout("🧠", "purple_background", decisionsText),
    buildHeading2("Top issues"),
    buildCallout("🚨", "red_background", topIssuesText),
    buildHeading2("Risk type breakdown"),
    buildCallout("🧷", "gray_background", riskTypeText),
    buildHeading2("Dependency and milestone impact"),
    buildCallout("🔗", "yellow_background", dependencySummaryText),
    buildHeading2("How upstream/downstream impact works"),
    buildCallout("🧭", "yellow_background", dependencyBasicsText),
    buildHeading2("Top pressure workstreams"),
    buildCallout("📊", "gray_background", wsText),
    buildParagraph("Use the full task scheduling dashboard for details."),
    {
      object: "block",
      type: "paragraph",
      paragraph: {
        rich_text: [
          rt("Open task scheduling dashboard", {
            url: `https://www.notion.so/${String(SCHEDULE_DASHBOARD_PAGE_ID).replace(/-/g, "")}`,
            annotations: { bold: true }
          })
        ]
      }
    }
  ]);

  await replacePageContent(pmoPage.id, [
    {
      object: "block",
      type: "heading_1",
      heading_1: { rich_text: [rt("PMO — Schedule Working View")] }
    },
    buildParagraph(`Updated ${today}`),
    buildHeading2("What this tells us right now"),
    buildCallout("🧠", "yellow_background", soWhatText),
    buildCallout("🗣️", "gray_background", summarySentence),
    buildHeading2("KPI movement vs previous snapshot"),
    buildCallout(
      "📈",
      "gray_background",
      previousSnapshot
        ? [
            `Reference date: ${previousSnapshot.date}`,
            `• Open tasks ${fmtSigned(deltas.open)}`,
            `• Slipped open ${fmtSigned(deltas.slippedOpen)}`,
            `• Due in 14 days ${fmtSigned(deltas.due14)}`,
            `• Overdue starts ${fmtSigned(deltas.overdueStarts)}`
          ].join("\n")
        : "No prior snapshot exists yet for trend comparison."
    ),
    buildHeading2("Execution watchlist"),
    buildCallout("⚠️", "orange_background", slippedText),
    buildHeading2("Top issues"),
    buildCallout("🚨", "red_background", topIssuesText),
    buildHeading2("Risk type breakdown"),
    buildCallout("🧷", "gray_background", riskTypeText),
    buildHeading2("Dependency change and milestone impact"),
    buildCallout("🔗", "yellow_background", dependencySummaryText),
    buildHeading2("How upstream/downstream impact works"),
    buildCallout("🧭", "yellow_background", dependencyBasicsText),
    buildHeading2("Top dependency changes"),
    buildCallout("🔄", "orange_background", dependencyChangeText),
    buildHeading2("Milestone impacts"),
    buildCallout("🏁", "blue_background", milestoneImpactText),
    buildHeading2("Overdue starts"),
    buildCallout("⏳", "red_background", overdueText),
    buildHeading2("Near-term due load (14d)"),
    buildCallout("📌", "blue_background", due14Text),
    buildHeading2("Top workstream pressure"),
    buildCallout("📊", "gray_background", wsText),
    buildParagraph("Use this page for weekly PMO working sessions; use steering view for executive forums.")
  ]);

  console.log(
    JSON.stringify(
      {
        scheduleDashboardUrl: `https://www.notion.so/${String(SCHEDULE_DASHBOARD_PAGE_ID).replace(/-/g, "")}`,
        steeringPageUrl: steeringPage.url,
        pmoPageUrl: pmoPage.url,
        snapshotDbUrl: snapshotDb.url,
        health: `${m.healthEmoji} ${m.healthLabel}`,
        kpis: {
          total: m.total,
          open: m.open,
          slippedOpen: m.slippedOpen,
          slipRateOpenPct: Number(m.slipRateOpen.toFixed(1)),
          due14: m.due14,
          overdueStarts: m.overdueStarts
        }
      },
      null,
      2
    )
  );
}

main().catch((error) => {
  console.error("refresh-schedule-dashboard failed:", error?.body || error?.message || error);
  process.exit(1);
});

/* eslint-disable no-console */
require("dotenv/config");

const { Client } = require("@notionhq/client");

const notion = new Client({ auth: process.env.NOTION_API_KEY });

const TASKS_DB_ID = process.env.NOTION_TASKS_DB_ID;
const RISKS_DB_ID = process.env.NOTION_RISKS_DB_ID || "357ae9be-8a60-8190-b720-c130c7104cf1";

const TASK_RISK_PROP = "Risks & Issues";
const RISK_TASK_PROP = "Linked Tasks";
const MAX_TASKS_PER_RECORD = Number(process.env.RISK_TASK_LINK_LIMIT || 50);
const MAX_GROUPS = Number(process.env.RISK_GROUP_LIMIT || 8);
const MAX_FAMILY_GROUPS = Number(process.env.RISK_FAMILY_GROUP_LIMIT || 6);
const CONCURRENCY = Number(process.env.RISK_SYNC_CONCURRENCY || 2);
const DELAY_MS = Number(process.env.RISK_SYNC_DELAY_MS || 150);

function sleep(ms) {
  return new Promise((resolve) => setTimeout(resolve, ms));
}

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

function daysDiff(current, baseline) {
  if (!current || !baseline) return null;
  const currentDate = new Date(current);
  const baselineDate = new Date(baseline);
  if (Number.isNaN(currentDate.getTime()) || Number.isNaN(baselineDate.getTime())) {
    return null;
  }
  return Math.round((currentDate - baselineDate) / (1000 * 60 * 60 * 24));
}

function relationArray(ids) {
  return ids.map((id) => ({ id }));
}

function richText(value) {
  if (!value) return [];
  return [{ type: "text", text: { content: String(value).slice(0, 1900) } }];
}

function titleText(value) {
  return [{ type: "text", text: { content: String(value).slice(0, 1900) } }];
}

function isRetryableError(error) {
  const status = error?.status || error?.body?.status;
  const code = error?.code || error?.body?.code;
  return (
    status === 429 ||
    status >= 500 ||
    code === "rate_limited" ||
    code === "request_timeout" ||
    code === "notionhq_client_request_timeout"
  );
}

async function withRetry(fn, label, maxRetries = 6) {
  for (let attempt = 1; attempt <= maxRetries; attempt += 1) {
    try {
      return await fn();
    } catch (error) {
      if (!isRetryableError(error) || attempt === maxRetries) {
        throw error;
      }

      const retryAfterSec = Number(error?.headers?.get?.("retry-after") || error?.body?.retry_after || 0);
      const delayMs = Math.max(retryAfterSec * 1000, Math.min(2500 * attempt, 15000));
      console.log(`Retrying ${label} after ${delayMs}ms (${attempt}/${maxRetries})`);
      await sleep(delayMs);
    }
  }
}

async function fetchAllPages(databaseId, label) {
  const pages = [];
  let cursor;
  do {
    const result = await withRetry(
      () => notion.databases.query({ database_id: databaseId, start_cursor: cursor, page_size: 100 }),
      `query-${label}`
    );
    pages.push(...result.results.filter((row) => row.object === "page"));
    cursor = result.has_more ? result.next_cursor : undefined;
  } while (cursor);
  return pages;
}

async function runWithConcurrency(items, handler, label) {
  let index = 0;
  let completed = 0;

  async function worker() {
    while (true) {
      const current = index;
      if (current >= items.length) return;
      index += 1;
      if (DELAY_MS > 0) {
        await sleep(DELAY_MS);
      }
      await handler(items[current], current);
      completed += 1;
      if (completed % 50 === 0 || completed === items.length) {
        console.log(`${label}: ${completed}/${items.length}`);
      }
    }
  }

  await Promise.all(
    Array.from({ length: Math.min(CONCURRENCY, items.length || 1) }, () => worker())
  );
}

function priorityForGroup(group) {
  if (group.count >= 200 || group.maxSlip >= 150) return "Critical";
  if (group.count >= 50 || group.maxSlip >= 60) return "High";
  return "Medium";
}

function buildWorkstreamKey(workstream) {
  return `schedule-slippage::${workstream.toLowerCase()}`;
}

function buildFamilyKey(family) {
  return `delivery-theme::${family.toLowerCase()}`;
}

function classifyTaskFamily(taskName) {
  const normalized = String(taskName || "").toLowerCase();
  if (/load|simulate load|migration/.test(normalized)) return "Data Load / Migration";
  if (/extract|file staged|pre load|post load/.test(normalized)) return "Data Readiness";
  if (/defect|mitigation|retest/.test(normalized)) return "Defect Mitigation";
  if (/test|testing|uat|smoke|trr/.test(normalized)) return "Testing Readiness";
  if (/transport|release|deploy|cutover|close client/.test(normalized)) {
    return "Deployment / Transport";
  }
  if (/security|role/.test(normalized)) return "Security / Access";
  return "Other";
}

function buildRiskRecordSpec(group) {
  const isUnassigned = group.workstream === "(Unassigned)";
  const type = isUnassigned ? "Risk" : "Issue";
  const impact = isUnassigned ? "Resource" : "Schedule";
  const priority = priorityForGroup(group);
  const headline = isUnassigned
    ? "Unassigned workstream schedule slippage"
    : `${group.workstream} schedule slippage`;

  const sampleText = group.tasks
    .slice(0, 5)
    .map((task) => `${task.uid}: ${task.name} (${task.slipDays}d slip)`)
    .join("; ");

  const description = [
    `${group.count} incomplete tasks in workstream \"${group.workstream}\" are finishing later than baseline.`,
    `Maximum observed slip: ${group.maxSlip} days.`,
    sampleText ? `Sample tasks: ${sampleText}.` : ""
  ]
    .filter(Boolean)
    .join(" ");

  const mitigation = isUnassigned
    ? "Assign workstream ownership, review dashboard slippage views, and resolve resource accountability gaps."
    : "Review linked slipped tasks, confirm recovery plan dates, and align owners through dashboard and project status reviews.";

  return {
    key: buildWorkstreamKey(group.workstream),
    title: headline,
    type,
    impact,
    priority,
    description,
    mitigation,
    taskIds: group.tasks.slice(0, MAX_TASKS_PER_RECORD).map((task) => task.pageId)
  };
}

function buildFamilyRiskRecordSpec(group) {
  const familyRules = {
    "Data Load / Migration": {
      title: "Cross-workstream data load slippage",
      type: "Issue",
      impact: "Schedule",
      mitigation:
        "Review linked load and migration tasks, rebalance sequence dependencies, and escalate workstream blockers affecting load execution."
    },
    "Data Readiness": {
      title: "Data extract and staging readiness delays",
      type: "Issue",
      impact: "Schedule",
      mitigation:
        "Stabilize upstream extracts, staged files, and pre/post-load validation steps before the next dashboard review."
    },
    "Defect Mitigation": {
      title: "Defect mitigation backlog across workstreams",
      type: "Issue",
      impact: "Quality",
      mitigation:
        "Prioritize defect closure, align retest windows, and confirm recovery actions for high-slip defect mitigation activities."
    },
    "Testing Readiness": {
      title: "Testing readiness and TRR risk",
      type: "Risk",
      impact: "Quality",
      mitigation:
        "Review TRR, UAT, and smoke test readiness gates and assign owners before slippage converts into additional issues."
    },
    "Deployment / Transport": {
      title: "Deployment and transport readiness risk",
      type: "Risk",
      impact: "Schedule",
      mitigation:
        "Confirm transport approvals, deployment sequence readiness, and cutover prerequisites through the project status forum."
    },
    "Security / Access": {
      title: "Security and access readiness risk",
      type: "Risk",
      impact: "Resource",
      mitigation:
        "Resolve security role, tester access, and authorization dependencies before they delay later-cycle activities."
    }
  };

  const rule = familyRules[group.family];
  if (!rule) return null;

  const sampleText = group.tasks
    .slice(0, 5)
    .map((task) => `${task.uid}: ${task.name} [${task.workstream}] (${task.slipDays}d slip)`)
    .join("; ");

  return {
    key: buildFamilyKey(group.family),
    title: rule.title,
    type: rule.type,
    impact: rule.impact,
    priority: priorityForGroup(group),
    description: [
      `${group.count} slipped tasks align to the delivery theme \"${group.family}\" across multiple workstreams.`,
      `Maximum observed slip: ${group.maxSlip} days.`,
      sampleText ? `Sample tasks: ${sampleText}.` : ""
    ]
      .filter(Boolean)
      .join(" "),
    mitigation: rule.mitigation,
    owner: group.family,
    taskIds: group.tasks.slice(0, MAX_TASKS_PER_RECORD).map((task) => task.pageId)
  };
}

function pickTargetGroups(tasks) {
  const groups = new Map();

  for (const task of tasks) {
    const props = task.properties || {};
    const finish = dateValue(props.Finish);
    const baseline = dateValue(props["Baseline Finish"]) || dateValue(props["Baseline10 Finish"]);
    const percentComplete = numberValue(props["% Complete"]) ?? 0;
    const slipDays = daysDiff(finish, baseline);
    if (!(slipDays != null && slipDays > 0 && percentComplete < 100)) continue;

    const workstream = rich(props.Workstream).trim() || "(Unassigned)";
    const group = groups.get(workstream) || {
      workstream,
      count: 0,
      maxSlip: 0,
      tasks: []
    };

    group.count += 1;
    group.maxSlip = Math.max(group.maxSlip, slipDays);
    group.tasks.push({
      pageId: task.id,
      uid: props["Unique ID"]?.number ?? null,
      name: title(props["Task Name"]),
      slipDays
    });
    groups.set(workstream, group);
  }

  const ranked = [...groups.values()]
    .map((group) => ({
      ...group,
      tasks: group.tasks.sort((a, b) => b.slipDays - a.slipDays || (a.uid ?? 0) - (b.uid ?? 0))
    }))
    .filter((group) => group.count >= 20 || group.workstream === "(Unassigned)")
    .sort((a, b) => b.count - a.count || b.maxSlip - a.maxSlip)
    .slice(0, MAX_GROUPS);

  return ranked;
}

function pickFamilyGroups(tasks) {
  const groups = new Map();

  for (const task of tasks) {
    const props = task.properties || {};
    const finish = dateValue(props.Finish);
    const baseline = dateValue(props["Baseline Finish"]) || dateValue(props["Baseline10 Finish"]);
    const percentComplete = numberValue(props["% Complete"]) ?? 0;
    const slipDays = daysDiff(finish, baseline);
    if (!(slipDays != null && slipDays > 0 && percentComplete < 100)) continue;

    const taskName = title(props["Task Name"]);
    const family = classifyTaskFamily(taskName);
    if (family === "Other") continue;

    const workstream = rich(props.Workstream).trim() || "(Unassigned)";
    const group = groups.get(family) || {
      family,
      count: 0,
      maxSlip: 0,
      tasks: []
    };

    group.count += 1;
    group.maxSlip = Math.max(group.maxSlip, slipDays);
    group.tasks.push({
      pageId: task.id,
      uid: props["Unique ID"]?.number ?? null,
      name: taskName,
      workstream,
      slipDays
    });
    groups.set(family, group);
  }

  return [...groups.values()]
    .map((group) => ({
      ...group,
      tasks: group.tasks.sort((a, b) => b.slipDays - a.slipDays || (a.uid ?? 0) - (b.uid ?? 0))
    }))
    .filter((group) => group.count >= 15)
    .sort((a, b) => b.count - a.count || b.maxSlip - a.maxSlip)
    .slice(0, MAX_FAMILY_GROUPS);
}

async function main() {
  if (!process.env.NOTION_API_KEY || !TASKS_DB_ID) {
    throw new Error("NOTION_API_KEY and NOTION_TASKS_DB_ID are required.");
  }

  const [taskDb, riskDb, tasks, existingRiskRows] = await Promise.all([
    withRetry(() => notion.databases.retrieve({ database_id: TASKS_DB_ID }), "retrieve-tasks-db"),
    withRetry(() => notion.databases.retrieve({ database_id: RISKS_DB_ID }), "retrieve-risks-db"),
    fetchAllPages(TASKS_DB_ID, "tasks"),
    fetchAllPages(RISKS_DB_ID, "risks")
  ]);

  if (!taskDb.properties?.[TASK_RISK_PROP]) {
    throw new Error(`Tasks DB is missing relation property \"${TASK_RISK_PROP}\".`);
  }
  if (!riskDb.properties?.[RISK_TASK_PROP]) {
    throw new Error(`Risks DB is missing relation property \"${RISK_TASK_PROP}\".`);
  }

  const selectedGroups = pickTargetGroups(tasks);
  const familyGroups = pickFamilyGroups(tasks);
  const specs = [
    ...selectedGroups.map(buildRiskRecordSpec),
    ...familyGroups.map(buildFamilyRiskRecordSpec).filter(Boolean)
  ];
  const existingByKey = new Map();
  const existingByTitle = new Map();

  for (const row of existingRiskRows) {
    const rowTitle = title(row.properties?.["Risk/Issue Name"]);
    if (rowTitle && !existingByTitle.has(rowTitle)) {
      existingByTitle.set(rowTitle, row);
    }

    const workstreamKey = buildWorkstreamKey(
      rowTitle
        .replace(/ schedule slippage$/i, "")
        .replace(/^Unassigned workstream schedule slippage$/i, "(Unassigned)")
    );
    if (!existingByKey.has(workstreamKey)) {
      existingByKey.set(workstreamKey, row);
    }

    const familyMap = {
      "Cross-workstream data load slippage": buildFamilyKey("Data Load / Migration"),
      "Data extract and staging readiness delays": buildFamilyKey("Data Readiness"),
      "Defect mitigation backlog across workstreams": buildFamilyKey("Defect Mitigation"),
      "Testing readiness and TRR risk": buildFamilyKey("Testing Readiness"),
      "Deployment and transport readiness risk": buildFamilyKey("Deployment / Transport"),
      "Security and access readiness risk": buildFamilyKey("Security / Access")
    };
    if (familyMap[rowTitle]) {
      if (!existingByKey.has(familyMap[rowTitle])) {
        existingByKey.set(familyMap[rowTitle], row);
      }
    }
  }

  const riskIdsByTaskId = new Map();
  let created = 0;
  let updated = 0;

  for (const spec of specs) {
    const properties = {
      "Risk/Issue Name": { title: titleText(spec.title) },
      Type: { select: { name: spec.type } },
      Status: { status: { name: "Open" } },
      Priority: { select: { name: spec.priority } },
      Impact: { select: { name: spec.impact } },
      Owner: { rich_text: richText(spec.owner || spec.title.replace(/ schedule slippage$/i, "")) },
      Description: { rich_text: richText(spec.description) },
      "Mitigation / Action": { rich_text: richText(spec.mitigation) },
      [RISK_TASK_PROP]: { relation: relationArray(spec.taskIds) }
    };

    const existing = existingByTitle.get(spec.title) || existingByKey.get(spec.key);
    if (existing) {
      await withRetry(
        () => notion.pages.update({ page_id: existing.id, properties }),
        `update-risk-${spec.key}`
      );
      updated += 1;
      for (const taskId of spec.taskIds) {
        if (!riskIdsByTaskId.has(taskId)) riskIdsByTaskId.set(taskId, new Set());
        riskIdsByTaskId.get(taskId).add(existing.id);
      }
    } else {
      const createdRow = await withRetry(
        () => notion.pages.create({ parent: { database_id: RISKS_DB_ID }, properties }),
        `create-risk-${spec.key}`
      );
      created += 1;
      existingByKey.set(spec.key, createdRow);
      for (const taskId of spec.taskIds) {
        if (!riskIdsByTaskId.has(taskId)) riskIdsByTaskId.set(taskId, new Set());
        riskIdsByTaskId.get(taskId).add(createdRow.id);
      }
    }
  }

  const taskPagesToUpdate = tasks.filter((task) => riskIdsByTaskId.has(task.id));
  let taskRelationsUpdated = 0;

  await runWithConcurrency(
    taskPagesToUpdate,
    async (task) => {
      const existingIds = (task.properties?.[TASK_RISK_PROP]?.relation || []).map((item) => item.id);
      const merged = [...new Set([...existingIds, ...riskIdsByTaskId.get(task.id)])];
      if (merged.length === existingIds.length) return;

      await withRetry(
        () =>
          notion.pages.update({
            page_id: task.id,
            properties: {
              [TASK_RISK_PROP]: { relation: relationArray(merged) }
            }
          }),
        `update-task-risk-${task.id}`
      );
      taskRelationsUpdated += 1;
    },
    "Task risk relation sync"
  );

  console.log("\nRisk/issue sync complete:");
  console.log(`  Groups evaluated: ${selectedGroups.length}`);
  console.log(`  Family themes evaluated: ${familyGroups.length}`);
  console.log(`  Risk/issue rows created: ${created}`);
  console.log(`  Risk/issue rows updated: ${updated}`);
  console.log(`  Tasks linked to risks/issues: ${taskPagesToUpdate.length}`);
  console.log(`  Task-side relations updated: ${taskRelationsUpdated}`);
}

main().catch((error) => {
  console.error("Risk/issue sync failed:", error?.body || error?.message || error);
  process.exit(1);
});

/* eslint-disable no-console */
require("dotenv/config");

const { Client } = require("@notionhq/client");

const notion = new Client({ auth: process.env.NOTION_API_KEY });

const TASKS_DB_ID = process.env.NOTION_TASKS_DB_ID;
const PROJECTS_DB_ID = process.env.NOTION_PROJECTS_DB_ID;
const AMETEK_PROJECT_PAGE_ID = process.env.NOTION_AMETEK_PROJECT_PAGE_ID;
const RISKS_DB_ID_ENV = process.env.NOTION_RISKS_DB_ID;
const DASHBOARD_PAGE_ID_ENV = process.env.NOTION_DASHBOARD_PAGE_ID;
const PROJECT_STATUS_PAGE_ID_ENV = process.env.NOTION_PROJECT_STATUS_PAGE_ID;

const TASK_RISK_REL_PROP = "Risks & Issues";
const TASK_PROJECT_REL_PROP = "Projects";
const TASK_DASHBOARD_URL_PROP = "Dashboard Link";
const TASK_STATUS_URL_PROP = "Project Status Link";
const TASK_RISKS_URL_PROP = "Risks & Issues Link";

const CONCURRENCY = Number(process.env.CONTEXT_LINK_CONCURRENCY || 2);
const DELAY_MS = Number(process.env.CONTEXT_LINK_DELAY_MS || 150);
const TIMEOUT_MS = Number(process.env.CONTEXT_LINK_TIMEOUT_MS || 45000);

function sleep(ms) {
  return new Promise((resolve) => setTimeout(resolve, ms));
}

function assertEnv() {
  if (!process.env.NOTION_API_KEY) {
    throw new Error("NOTION_API_KEY is required in .env");
  }
  if (!TASKS_DB_ID) {
    throw new Error("NOTION_TASKS_DB_ID is required in .env");
  }
  if (!Number.isFinite(CONCURRENCY) || CONCURRENCY < 1) {
    throw new Error("CONTEXT_LINK_CONCURRENCY must be a positive integer.");
  }
}

function getText(parts) {
  return (parts || []).map((p) => p.plain_text).join("");
}

function getTitleFromPage(page) {
  const props = page.properties || {};
  const titleProp = Object.values(props).find((p) => p?.type === "title");
  if (!titleProp || !titleProp.title?.length) return "";
  return getText(titleProp.title).trim();
}

function isRetryableError(error) {
  const status = error?.status || error?.body?.status;
  const code = error?.code || error?.body?.code;
  return status === 429 || status >= 500 || code === "rate_limited" || code === "request_timeout";
}

async function withTimeout(promise, ms, label) {
  let timeoutHandle;
  const timeoutPromise = new Promise((_, reject) => {
    timeoutHandle = setTimeout(() => {
      const err = new Error(`Timeout after ${ms}ms: ${label}`);
      err.code = "request_timeout";
      reject(err);
    }, ms);
  });

  try {
    return await Promise.race([promise, timeoutPromise]);
  } finally {
    clearTimeout(timeoutHandle);
  }
}

async function withRetry(fn, label, maxRetries = 8) {
  let attempt = 0;

  while (true) {
    try {
      return await withTimeout(fn(), TIMEOUT_MS, label);
    } catch (error) {
      attempt += 1;
      if (!isRetryableError(error) || attempt > maxRetries) {
        throw error;
      }

      const retryAfterSec = Number(error?.headers?.get?.("retry-after") || error?.body?.retry_after);
      const retryAfterMs = Number.isFinite(retryAfterSec) ? retryAfterSec * 1000 : 0;
      const delayMs = Math.max(retryAfterMs, Math.min(2500 * attempt, 15000));
      console.log(`Retrying ${label} after ${delayMs}ms (attempt ${attempt}/${maxRetries})`);
      await sleep(delayMs);
    }
  }
}

async function runWithConcurrency(items, limit, handler, phaseLabel) {
  const startedAt = Date.now();
  let cursor = 0;
  let completed = 0;
  let nextLogAt = 200;

  async function worker() {
    while (true) {
      const index = cursor;
      if (index >= items.length) return;
      cursor += 1;

      if (DELAY_MS > 0) {
        await sleep(DELAY_MS);
      }

      await handler(items[index], index);
      completed += 1;

      if (completed >= nextLogAt || completed === items.length) {
        const seconds = Math.round((Date.now() - startedAt) / 1000);
        console.log(`${phaseLabel}: ${completed}/${items.length} (${seconds}s)`);
        nextLogAt += 200;
      }
    }
  }

  const workers = Array.from({ length: Math.min(limit, items.length) }, () => worker());
  await Promise.all(workers);
}

async function fetchAllPages(databaseId, label) {
  const pages = [];
  let cursor;

  do {
    const res = await withRetry(
      () => notion.databases.query({ database_id: databaseId, start_cursor: cursor, page_size: 100 }),
      `query-${label}`
    );

    pages.push(...res.results.filter((r) => r.object === "page"));
    cursor = res.has_more ? res.next_cursor : undefined;
  } while (cursor);

  return pages;
}

async function searchDatabaseByExactTitle(title) {
  let cursor;
  do {
    const res = await withRetry(
      () =>
        notion.search({
          query: title,
          filter: { value: "database", property: "object" },
          start_cursor: cursor,
          page_size: 100
        }),
      `search-db-${title}`
    );

    for (const item of res.results) {
      if (item.object !== "database") continue;
      const dbTitle = getText(item.title).trim().toLowerCase();
      if (dbTitle === title.trim().toLowerCase()) {
        return item;
      }
    }

    cursor = res.has_more ? res.next_cursor : undefined;
  } while (cursor);

  return null;
}

async function searchPageByTitleIncludes(partialTitle) {
  let cursor;
  const needle = partialTitle.trim().toLowerCase();

  do {
    const res = await withRetry(
      () => notion.search({ query: partialTitle, start_cursor: cursor, page_size: 100 }),
      `search-page-${partialTitle}`
    );

    for (const item of res.results) {
      if (item.object !== "page") continue;
      const titleProp = item.properties?.title;
      if (titleProp?.type !== "title") continue;
      const title = getText(titleProp.title).trim().toLowerCase();
      if (title.includes(needle)) {
        return item;
      }
    }

    cursor = res.has_more ? res.next_cursor : undefined;
  } while (cursor);

  return null;
}

function relationArray(ids) {
  return ids.map((id) => ({ id }));
}

async function resolveAmetekProjectPageId(projectsDbId) {
  if (AMETEK_PROJECT_PAGE_ID && AMETEK_PROJECT_PAGE_ID.trim()) {
    return AMETEK_PROJECT_PAGE_ID.trim();
  }

  if (!projectsDbId) return null;

  const projectRows = await fetchAllPages(projectsDbId, "projects");
  for (const row of projectRows) {
    const title = getTitleFromPage(row).toLowerCase();
    if (title.includes("ametek")) {
      return row.id;
    }
  }

  return null;
}

function findTaskRelationPropOnRiskDb(riskDbProps) {
  for (const [name, prop] of Object.entries(riskDbProps || {})) {
    if (prop?.type !== "relation") continue;
    if (prop.relation?.database_id === TASKS_DB_ID) {
      return name;
    }
  }
  return null;
}

async function ensureTaskProperties({ risksDbId, taskDbProperties }) {
  const addProps = {};

  if (risksDbId && !taskDbProperties[TASK_RISK_REL_PROP]) {
    addProps[TASK_RISK_REL_PROP] = {
      relation: { database_id: risksDbId, single_property: {} }
    };
  }

  if (!taskDbProperties[TASK_DASHBOARD_URL_PROP]) {
    addProps[TASK_DASHBOARD_URL_PROP] = { url: {} };
  }

  if (!taskDbProperties[TASK_STATUS_URL_PROP]) {
    addProps[TASK_STATUS_URL_PROP] = { url: {} };
  }

  if (!taskDbProperties[TASK_RISKS_URL_PROP]) {
    addProps[TASK_RISKS_URL_PROP] = { url: {} };
  }

  if (Object.keys(addProps).length === 0) return;

  await withRetry(
    () => notion.databases.update({ database_id: TASKS_DB_ID, properties: addProps }),
    "ensure-task-context-properties"
  );
}

async function main() {
  assertEnv();

  const taskDbBefore = await withRetry(
    () => notion.databases.retrieve({ database_id: TASKS_DB_ID }),
    "retrieve-task-db"
  );

  const risksDbId = RISKS_DB_ID_ENV?.trim() || (await searchDatabaseByExactTitle("Risks & Issues"))?.id || null;
  const dashboardPage = DASHBOARD_PAGE_ID_ENV?.trim()
    ? await withRetry(() => notion.pages.retrieve({ page_id: DASHBOARD_PAGE_ID_ENV.trim() }), "retrieve-dashboard-page")
    : await searchPageByTitleIncludes("Dashboard");
  const projectStatusPage = PROJECT_STATUS_PAGE_ID_ENV?.trim()
    ? await withRetry(() => notion.pages.retrieve({ page_id: PROJECT_STATUS_PAGE_ID_ENV.trim() }), "retrieve-project-status-page")
    : await searchPageByTitleIncludes("Project Status");

  await ensureTaskProperties({
    risksDbId,
    taskDbProperties: taskDbBefore.properties || {}
  });

  const taskDb = await withRetry(
    () => notion.databases.retrieve({ database_id: TASKS_DB_ID }),
    "retrieve-task-db-after-ensure"
  );
  const taskProps = taskDb.properties || {};

  const projectsDbId = taskProps[TASK_PROJECT_REL_PROP]?.relation?.database_id || PROJECTS_DB_ID || null;
  const ametekProjectId = await resolveAmetekProjectPageId(projectsDbId);

  let riskTaskPropName = null;
  let riskPages = [];
  if (risksDbId) {
    const riskDb = await withRetry(
      () => notion.databases.retrieve({ database_id: risksDbId }),
      "retrieve-risks-db"
    );
    riskTaskPropName = findTaskRelationPropOnRiskDb(riskDb.properties || {});
    if (riskTaskPropName) {
      riskPages = await fetchAllPages(risksDbId, "risks-db");
    }
  }

  const risksByTaskId = new Map();
  if (riskTaskPropName) {
    for (const riskPage of riskPages) {
      const rel = riskPage.properties?.[riskTaskPropName]?.relation || [];
      for (const taskRef of rel) {
        if (!risksByTaskId.has(taskRef.id)) risksByTaskId.set(taskRef.id, new Set());
        risksByTaskId.get(taskRef.id).add(riskPage.id);
      }
    }
  }

  const taskPages = await fetchAllPages(TASKS_DB_ID, "tasks-db");

  let updated = 0;
  let failed = 0;
  let projectLinked = 0;
  let riskLinked = 0;
  let dashboardLinked = 0;
  let statusLinked = 0;

  await runWithConcurrency(
    taskPages,
    CONCURRENCY,
    async (taskPage) => {
      const propsToUpdate = {};
      const taskId = taskPage.id;

      if (ametekProjectId && taskProps[TASK_PROJECT_REL_PROP]?.type === "relation") {
        const existing = (taskPage.properties?.[TASK_PROJECT_REL_PROP]?.relation || []).map((r) => r.id);
        if (!existing.includes(ametekProjectId)) {
          const merged = [...new Set([...existing, ametekProjectId])];
          propsToUpdate[TASK_PROJECT_REL_PROP] = { relation: relationArray(merged) };
          projectLinked += 1;
        }
      }

      if (taskProps[TASK_RISK_REL_PROP]?.type === "relation") {
        const existing = (taskPage.properties?.[TASK_RISK_REL_PROP]?.relation || []).map((r) => r.id);
        const fromRisks = [...(risksByTaskId.get(taskId) || new Set())];
        const merged = [...new Set([...existing, ...fromRisks])];
        if (merged.length !== existing.length) {
          propsToUpdate[TASK_RISK_REL_PROP] = { relation: relationArray(merged) };
          riskLinked += 1;
        }
      }

      const dashboardUrl = dashboardPage?.url || null;
      const statusUrl = projectStatusPage?.url || null;
      const risksDbUrl = risksDbId ? `https://www.notion.so/${String(risksDbId).replace(/-/g, "")}` : null;

      if (taskProps[TASK_DASHBOARD_URL_PROP]?.type === "url" && dashboardUrl) {
        const existingUrl = taskPage.properties?.[TASK_DASHBOARD_URL_PROP]?.url || null;
        if (existingUrl !== dashboardUrl) {
          propsToUpdate[TASK_DASHBOARD_URL_PROP] = { url: dashboardUrl };
          dashboardLinked += 1;
        }
      }

      if (taskProps[TASK_STATUS_URL_PROP]?.type === "url" && statusUrl) {
        const existingUrl = taskPage.properties?.[TASK_STATUS_URL_PROP]?.url || null;
        if (existingUrl !== statusUrl) {
          propsToUpdate[TASK_STATUS_URL_PROP] = { url: statusUrl };
          statusLinked += 1;
        }
      }

      if (taskProps[TASK_RISKS_URL_PROP]?.type === "url" && risksDbUrl) {
        const existingUrl = taskPage.properties?.[TASK_RISKS_URL_PROP]?.url || null;
        if (existingUrl !== risksDbUrl) {
          propsToUpdate[TASK_RISKS_URL_PROP] = { url: risksDbUrl };
        }
      }

      if (Object.keys(propsToUpdate).length === 0) return;

      try {
        await withRetry(
          () => notion.pages.update({ page_id: taskId, properties: propsToUpdate }),
          `update-task-context-${taskId}`
        );
        updated += 1;
      } catch (error) {
        failed += 1;
        console.log(`Failed update ${taskId}: ${error?.message || error}`);
      }
    },
    "Context link phase"
  );

  console.log("\nContext linking complete:");
  console.log(`  Tasks scanned: ${taskPages.length}`);
  console.log(`  Tasks updated: ${updated}`);
  console.log(`  Failed updates: ${failed}`);
  console.log(`  Added Projects links: ${projectLinked}`);
  console.log(`  Added Risks & Issues relation links: ${riskLinked}`);
  console.log(`  Added/updated Dashboard links: ${dashboardLinked}`);
  console.log(`  Added/updated Project Status links: ${statusLinked}`);
  console.log(`  Risks DB: ${risksDbId || "not found"}`);
  console.log(`  Dashboard page: ${dashboardPage?.url || "not found"}`);
  console.log(`  Project status page: ${projectStatusPage?.url || "not found"}`);
  console.log(`  Ametek project row: ${ametekProjectId || "not found"}`);
}

main().catch((error) => {
  console.error("Context linking failed:", error?.body || error?.message || error);
  process.exit(1);
});

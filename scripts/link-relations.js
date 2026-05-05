/* eslint-disable no-console */
require("dotenv/config");

const { Client } = require("@notionhq/client");

const notion = new Client({ auth: process.env.NOTION_API_KEY });

const TASKS_DB_ID = process.env.NOTION_TASKS_DB_ID;
const PROJECTS_DB_ID = process.env.NOTION_PROJECTS_DB_ID;
const PROJECT_PAGE_ID = process.env.NOTION_AMETEK_PROJECT_PAGE_ID;
const LINK_CONCURRENCY = Number(process.env.LINK_CONCURRENCY || 1);
const LINK_REQUEST_DELAY_MS = Number(process.env.LINK_REQUEST_DELAY_MS || 250);
const LINK_TIMEOUT_MS = Number(process.env.LINK_TIMEOUT_MS || 45000);

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
  if (!Number.isFinite(LINK_CONCURRENCY) || LINK_CONCURRENCY < 1) {
    throw new Error("LINK_CONCURRENCY must be a positive integer.");
  }
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

async function withRetry(fn, label, maxRetries = 6) {
  let attempt = 0;

  while (true) {
    try {
      return await withTimeout(fn(), LINK_TIMEOUT_MS, label);
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
  let nextLogAt = 250;

  async function worker() {
    while (true) {
      const index = cursor;
      if (index >= items.length) {
        return;
      }
      cursor += 1;

      if (LINK_REQUEST_DELAY_MS > 0) {
        await sleep(LINK_REQUEST_DELAY_MS);
      }

      await handler(items[index], index);
      completed += 1;

      if (completed >= nextLogAt || completed === items.length) {
        const seconds = Math.round((Date.now() - startedAt) / 1000);
        console.log(`${phaseLabel}: ${completed}/${items.length} (${seconds}s)`);
        nextLogAt += 250;
      }
    }
  }

  const workers = Array.from({ length: Math.min(limit, items.length) }, () => worker());
  await Promise.all(workers);
}

function getNumberProperty(page, propName) {
  const prop = page.properties[propName];
  if (!prop || prop.type !== "number") return null;
  return prop.number;
}

function getRichTextProperty(page, propName) {
  const prop = page.properties[propName];
  if (!prop || prop.type !== "rich_text") return "";
  return prop.rich_text.map((x) => x.plain_text).join("");
}

function parseIdList(value) {
  if (!value) return [];
  const matches = String(value).match(/\d+/g);
  if (!matches) return [];
  return [...new Set(matches.map(Number).filter(Number.isFinite))];
}

function relationArray(ids) {
  return ids.map((id) => ({ id }));
}

function parentHierarchyKey(key) {
  if (!key) return "";
  const parts = key.split(".").map((x) => x.trim()).filter(Boolean);
  if (parts.length <= 1) return "";
  return parts.slice(0, -1).join(".");
}

async function fetchAllTaskPages() {
  const pages = [];
  let cursor = undefined;

  do {
    const res = await withRetry(
      () => notion.databases.query({ database_id: TASKS_DB_ID, start_cursor: cursor, page_size: 100 }),
      "fetch-task-pages"
    );

    for (const r of res.results) {
      if (r.object === "page" && "properties" in r) {
        pages.push(r);
      }
    }

    cursor = res.has_more ? res.next_cursor : undefined;
  } while (cursor);

  return pages;
}

async function findAmetekProjectPageId() {
  if (PROJECT_PAGE_ID && PROJECT_PAGE_ID.trim()) {
    return PROJECT_PAGE_ID.trim();
  }
  if (!PROJECTS_DB_ID) {
    return undefined;
  }

  const res = await withRetry(
    () => notion.databases.query({ database_id: PROJECTS_DB_ID, page_size: 100 }),
    "find-project"
  );

  for (const row of res.results) {
    if (!(row.object === "page" && "properties" in row)) continue;
    const nameProp = row.properties.Name;
    if (!nameProp || nameProp.type !== "title") continue;
    const text = nameProp.title.map((x) => x.plain_text).join("").toLowerCase();
    if (text.includes("ametek")) {
      return row.id;
    }
  }

  return undefined;
}

async function main() {
  assertEnv();

  const pages = await fetchAllTaskPages();
  const pageIdByUniqueId = new Map();
  const pageIdByHierarchy = new Map();

  for (const page of pages) {
    const uniqueId = getNumberProperty(page, "Unique ID");
    if (uniqueId != null) {
      pageIdByUniqueId.set(uniqueId, page.id);
    }

    const hierarchy = getRichTextProperty(page, "Outline Number") || getRichTextProperty(page, "WBS");
    if (hierarchy) {
      pageIdByHierarchy.set(hierarchy, page.id);
    }
  }

  const projectPageId = await findAmetekProjectPageId();

  let parentLinked = 0;
  let dependencyLinked = 0;
  let projectLinked = 0;
  let unresolvedParents = 0;
  let unresolvedDependencies = 0;
  let failed = 0;

  await runWithConcurrency(
    pages,
    LINK_CONCURRENCY,
    async (page) => {
      const properties = page.properties;
      const uniqueId = getNumberProperty(page, "Unique ID");
      if (uniqueId == null) return;

      const hierarchy = getRichTextProperty(page, "Outline Number") || getRichTextProperty(page, "WBS");
      const parentKey = parentHierarchyKey(hierarchy);
      let parentId = null;
      if (parentKey) {
        parentId = pageIdByHierarchy.get(parentKey) || null;
        if (parentId) parentLinked += 1;
        else unresolvedParents += 1;
      }

      const predecessorsText =
        getRichTextProperty(page, "Unique ID Predecessors") || getRichTextProperty(page, "Predecessors");
      const successorsText = getRichTextProperty(page, "Unique ID Successors");

      const predecessorUids = parseIdList(predecessorsText);
      const successorUids = parseIdList(successorsText);

      const predecessorIds = predecessorUids
        .map((id) => pageIdByUniqueId.get(id))
        .filter(Boolean);
      const successorIds = successorUids
        .map((id) => pageIdByUniqueId.get(id))
        .filter(Boolean);

      if (predecessorIds.length < predecessorUids.length) {
        unresolvedDependencies += predecessorUids.length - predecessorIds.length;
      }
      if (successorIds.length < successorUids.length) {
        unresolvedDependencies += successorUids.length - successorIds.length;
      }

      dependencyLinked += predecessorIds.length + successorIds.length;

      const relationProps = {
        "Parent Task": { relation: parentId ? relationArray([parentId]) : [] },
        "Predecessor Tasks": { relation: relationArray([...new Set(predecessorIds)]) },
        "Successor Tasks": { relation: relationArray([...new Set(successorIds)]) }
      };

      if (projectPageId) {
        relationProps.Projects = { relation: relationArray([projectPageId]) };
        projectLinked += 1;
      }

      try {
        await withRetry(
          () => notion.pages.update({ page_id: page.id, properties: relationProps }),
          `link-relations-${uniqueId}`
        );
      } catch (error) {
        failed += 1;
        console.log(`Failed row Unique ID ${uniqueId}: ${error?.message || error}`);
      }
    },
    "Link phase"
  );

  console.log("\nRelation linking complete:");
  console.log(`  Tasks processed: ${pages.length}`);
  console.log(`  Parent links set: ${parentLinked}`);
  console.log(`  Dependency links set: ${dependencyLinked}`);
  console.log(`  Project links set: ${projectLinked}`);
  console.log(`  Unresolved parent links: ${unresolvedParents}`);
  console.log(`  Unresolved dependency refs: ${unresolvedDependencies}`);
  console.log(`  Failed row updates: ${failed}`);
}

main().catch((error) => {
  console.error("Linking failed:", error?.body || error?.message || error);
  process.exit(1);
});

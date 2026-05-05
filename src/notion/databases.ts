import type {
  DatabaseObjectResponse,
  PageObjectResponse,
  QueryDatabaseResponse
} from "@notionhq/client/build/src/api-endpoints";
import { env } from "../config/env";
import { notion } from "./client";

const notionStatusProperty = env.NOTION_STATUS_PROPERTY ?? "Status";
const notionProgressProperty = env.NOTION_PROGRESS_PROPERTY ?? "Progress";

export type AccessibleDatabase = {
  id: string;
  title: string;
  url: string;
};

export type RiskRow = {
  pageId: string;
  title: string;
  status?: string;
  progress?: number;
};

function getPlainText(parts: Array<{ plain_text: string }>): string {
  return parts.map((part) => part.plain_text).join("");
}

function getDatabaseTitle(database: DatabaseObjectResponse): string {
  const title = getPlainText(database.title);
  return title || "(untitled database)";
}

function getTitle(properties: Record<string, any>): string {
  const titleProp = Object.values(properties).find(
    (prop: any) => prop?.type === "title"
  ) as any;

  if (!titleProp?.title?.length) {
    return "(untitled)";
  }

  return getPlainText(titleProp.title);
}

function extractStatus(properties: Record<string, any>): string | undefined {
  const prop = properties[notionStatusProperty] as any;

  if (!prop) return undefined;
  if (prop.type === "status") return prop.status?.name;
  if (prop.type === "select") return prop.select?.name;

  return undefined;
}

function extractProgress(properties: Record<string, any>): number | undefined {
  const prop = properties[notionProgressProperty] as any;

  if (!prop) return undefined;
  if (prop.type === "number") return prop.number ?? undefined;

  return undefined;
}

export async function listAccessibleDatabases(
  searchQuery?: string
): Promise<AccessibleDatabase[]> {
  const databases: AccessibleDatabase[] = [];
  let cursor: string | undefined;

  do {
    const response = await notion.search({
      query: searchQuery,
      filter: {
        property: "object",
        value: "database"
      },
      start_cursor: cursor,
      page_size: 100
    });

    const pageResults = response.results.filter(
      (result): result is DatabaseObjectResponse => result.object === "database"
    );

    databases.push(
      ...pageResults.map((database) => ({
        id: database.id,
        title: getDatabaseTitle(database),
        url: database.url
      }))
    );

    cursor = response.has_more ? response.next_cursor ?? undefined : undefined;
  } while (cursor);

  return databases;
}

export async function resolveRiskDatabase(): Promise<AccessibleDatabase> {
  if (env.NOTION_DATABASE_ID) {
    return {
      id: env.NOTION_DATABASE_ID,
      title: env.NOTION_DATABASE_NAME ?? "Configured database",
      url: ""
    };
  }

  const databases = await listAccessibleDatabases(env.NOTION_DATABASE_NAME);

  if (databases.length === 0) {
    throw new Error(
      "No shared Notion databases were found for this integration token. Make sure the database is shared with the integration, or set NOTION_DATABASE_ID manually."
    );
  }

  if (env.NOTION_DATABASE_NAME) {
    const exactMatches = databases.filter(
      (database) =>
        database.title.toLowerCase() === env.NOTION_DATABASE_NAME?.toLowerCase()
    );

    if (exactMatches.length === 1) {
      return exactMatches[0];
    }

    if (exactMatches.length > 1) {
      const options = exactMatches
        .map((database) => `- ${database.title} (${database.id})`)
        .join("\n");
      throw new Error(
        `Multiple databases matched NOTION_DATABASE_NAME exactly. Set NOTION_DATABASE_ID to disambiguate.\n${options}`
      );
    }
  }

  if (databases.length === 1) {
    return databases[0];
  }

  const options = databases
    .slice(0, 20)
    .map((database) => `- ${database.title} (${database.id})`)
    .join("\n");

  throw new Error(
    `Multiple shared Notion databases are available for this token. Set NOTION_DATABASE_NAME or NOTION_DATABASE_ID.\n${options}`
  );
}

export async function readRiskRows(limit = 25): Promise<RiskRow[]> {
  const database = await resolveRiskDatabase();
  const response = (await notion.databases.query({
    database_id: database.id,
    page_size: limit
  })) as QueryDatabaseResponse;

  return response.results
    .filter(
      (r): r is PageObjectResponse =>
        r.object === "page" && "properties" in r
    )
    .map((page) => {
      const properties = page.properties as Record<string, any>;
      return {
        pageId: page.id,
        title: getTitle(properties),
        status: extractStatus(properties),
        progress: extractProgress(properties)
      };
    });
}

export async function updateRiskRow(
  pageId: string,
  updates: {
    status?: string;
    progress?: number;
  }
): Promise<void> {
  const properties: Record<string, any> = {};

  if (typeof updates.status === "string") {
    properties[notionStatusProperty] = {
      status: {
        name: updates.status
      }
    };
  }

  if (typeof updates.progress === "number") {
    properties[notionProgressProperty] = {
      number: updates.progress
    };
  }

  if (Object.keys(properties).length === 0) {
    throw new Error("No updates provided. Pass status and/or progress.");
  }

  await notion.pages.update({
    page_id: pageId,
    properties
  });
}

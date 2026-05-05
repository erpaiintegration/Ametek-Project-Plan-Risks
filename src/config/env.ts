import "dotenv/config";
import { z } from "zod";

const str = z.string().trim();
const optStr = str.optional().transform((v) => (v && v.length > 0 ? v : undefined));

const envSchema = z.object({
  NOTION_API_KEY: str.min(1, "NOTION_API_KEY is required"),
  NOTION_TASKS_DB_ID: str.min(1, "NOTION_TASKS_DB_ID is required"),
  NOTION_PROJECTS_DB_ID: optStr,
  NOTION_AMETEK_PROJECT_PAGE_ID: optStr,
  NOTES_OWNER: z.enum(["msp", "notion"]).default("notion"),
  NOTION_DATABASE_ID: optStr,
  NOTION_DATABASE_NAME: optStr,
  NOTION_STATUS_PROPERTY: optStr,
  NOTION_PROGRESS_PROPERTY: optStr
});

export const env = envSchema.parse(process.env);

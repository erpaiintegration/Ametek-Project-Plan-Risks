import {
  readRiskRows,
  resolveRiskDatabase,
  updateRiskRow
} from "./notion/databases";

async function main() {
  const database = await resolveRiskDatabase();
  console.log(`Using Notion database: ${database.title} (${database.id})`);
  console.log("Fetching Ametek risk rows from Notion...");
  const rows = await readRiskRows(10);

  if (rows.length === 0) {
    console.log("No rows found in the resolved Notion database.");
    return;
  }

  console.table(
    rows.map((r) => ({
      pageId: r.pageId,
      title: r.title,
      status: r.status ?? "",
      progress: r.progress ?? ""
    }))
  );

  // Demo update flow (safe default: only runs when env var is set)
  // Set DEMO_UPDATE_PAGE_ID in your shell/env to test updates intentionally.
  const demoPageId = process.env.DEMO_UPDATE_PAGE_ID;
  if (demoPageId) {
    console.log(`Applying demo update to page ${demoPageId} ...`);
    await updateRiskRow(demoPageId, {
      status: "In Progress",
      progress: 50
    });
    console.log("Demo update complete.");
  }
}

main().catch((error) => {
  console.error("App failed:", error);
  process.exit(1);
});

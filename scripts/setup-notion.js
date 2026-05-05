/**
 * setup-notion.js
 *
 * Creates in Notion (under the Ametek SAP S4 Implementation Project page):
 *   1. Risks & Issues database
 *   2. Management Dashboard page with linked views
 *
 * Run: node scripts/setup-notion.js
 */
require("dotenv").config();
const { Client } = require("@notionhq/client");

const notion = new Client({ auth: process.env.NOTION_API_KEY });

const AMETEK_PAGE_ID = "62bae9be-8a60-8398-b35a-81cf9997daa0";
const TASKS_DB_ID = "c15ae9be-8a60-83ee-8767-0152fa68585a";

// ─── 1. Create Risks & Issues DB ────────────────────────────────────────────

async function createRisksDatabase() {
  console.log("Creating Risks & Issues database...");

  const db = await notion.databases.create({
    parent: { type: "page_id", page_id: AMETEK_PAGE_ID },
    title: [{ type: "text", text: { content: "Risks & Issues" } }],
    icon: { type: "emoji", emoji: "⚠️" },
    properties: {
      "Risk/Issue Name": { title: {} },

      Type: {
        select: {
          options: [
            { name: "Risk", color: "yellow" },
            { name: "Issue", color: "red" },
            { name: "Conflict", color: "orange" },
            { name: "Dependency", color: "blue" }
          ]
        }
      },

      Status: {
        status: {
          options: [
            { name: "Open", color: "red" },
            { name: "In Progress", color: "yellow" },
            { name: "Resolved", color: "green" },
            { name: "Closed", color: "gray" }
          ],
          groups: [
            { name: "To-do", color: "gray", option_ids: [] },
            { name: "In progress", color: "blue", option_ids: [] },
            { name: "Complete", color: "green", option_ids: [] }
          ]
        }
      },

      Priority: {
        select: {
          options: [
            { name: "Critical", color: "red" },
            { name: "High", color: "orange" },
            { name: "Medium", color: "yellow" },
            { name: "Low", color: "green" }
          ]
        }
      },

      Impact: {
        select: {
          options: [
            { name: "Schedule", color: "red" },
            { name: "Budget", color: "orange" },
            { name: "Scope", color: "yellow" },
            { name: "Resource", color: "blue" },
            { name: "Quality", color: "purple" }
          ]
        }
      },

      "Linked Tasks": {
        relation: {
          database_id: TASKS_DB_ID,
          single_property: {}
        }
      },

      Owner: { rich_text: {} },
      "Due Date": { date: {} },
      "Raised Date": { date: {} },
      "Resolved Date": { date: {} },
      Description: { rich_text: {} },
      "Mitigation / Action": { rich_text: {} }
    }
  });

  console.log(`✅ Risks & Issues DB created: ${db.id}`);
  return db.id;
}

// ─── 2. Create Management Dashboard page ────────────────────────────────────

function dbLink(label) {
  // Notion API doesn't allow linked_to_page inside children blocks.
  // We use a callout as a placeholder the user can replace with a linked view.
  return {
    object: "block",
    type: "callout",
    callout: {
      icon: { type: "emoji", emoji: "🔗" },
      rich_text: [
        {
          type: "text",
          text: { content: `👉 Add a linked database view here: "${label}"` },
          annotations: { bold: true }
        },
        {
          type: "text",
          text: {
            content:
              "\nIn Notion: click the + below this block → Linked view of database → pick the database → apply the filter described above."
          },
          annotations: { italic: true, color: "gray" }
        }
      ],
      color: "blue_background"
    }
  };
}

async function createDashboardPage(risksDbId) {
  console.log("Creating Management Dashboard page...");

  const page = await notion.pages.create({
    parent: { type: "page_id", page_id: AMETEK_PAGE_ID },
    icon: { type: "emoji", emoji: "📊" },
    properties: {
      title: {
        title: [{ type: "text", text: { content: "Ametek Project Status — Management View" } }]
      }
    },
    children: [
      // Header
      {
        object: "block",
        type: "heading_1",
        heading_1: {
          rich_text: [{ type: "text", text: { content: "Ametek SAP S4 — Project Status Dashboard" } }]
        }
      },
      {
        object: "block",
        type: "paragraph",
        paragraph: {
          rich_text: [
            { type: "text", text: { content: "Last updated from MS Project: " } },
            { type: "text", text: { content: "run import to refresh" }, annotations: { italic: true, color: "gray" } }
          ]
        }
      },
      { object: "block", type: "divider", divider: {} },

      // ── Section 1: Critical Path ──
      {
        object: "block",
        type: "heading_2",
        heading_2: {
          rich_text: [{ type: "text", text: { content: "🔴  Critical Path Tasks" } }]
        }
      },
      {
        object: "block",
        type: "paragraph",
        paragraph: {
          rich_text: [{ type: "text", text: { content: "Tasks on the critical path that are not yet 100% complete. Any slip here moves the go-live date." } }]
        }
      },
      dbLink("Tasks — filter: Critical = ✓ AND % Complete < 100"),

      { object: "block", type: "divider", divider: {} },

      // ── Section 2: Slipping Tasks ──
      {
        object: "block",
        type: "heading_2",
        heading_2: {
          rich_text: [{ type: "text", text: { content: "⚠️  Slipping Tasks (Finish > Baseline Finish)" } }]
        }
      },
      {
        object: "block",
        type: "paragraph",
        paragraph: {
          rich_text: [{ type: "text", text: { content: "Tasks where the current finish date is later than the original baseline. These tasks have drifted from the plan." } }]
        }
      },
      dbLink("Tasks — filter: Finish is after Baseline Finish AND % Complete < 100"),

      { object: "block", type: "divider", divider: {} },

      // ── Section 3: Milestones ──
      {
        object: "block",
        type: "heading_2",
        heading_2: {
          rich_text: [{ type: "text", text: { content: "🏁  Upcoming Milestones (Next 60 Days)" } }]
        }
      },
      {
        object: "block",
        type: "paragraph",
        paragraph: {
          rich_text: [{ type: "text", text: { content: "Key milestone gates coming up. These are your go/no-go decision points." } }]
        }
      },
      dbLink("Tasks — filter: Milestone = ✓, sort Finish ascending"),

      { object: "block", type: "divider", divider: {} },

      // ── Section 4: Open Risks & Issues ──
      {
        object: "block",
        type: "heading_2",
        heading_2: {
          rich_text: [{ type: "text", text: { content: "🔥  Open Risks & Issues" } }]
        }
      },
      {
        object: "block",
        type: "paragraph",
        paragraph: {
          rich_text: [{ type: "text", text: { content: "Active risks and issues requiring owner action or management decision. Sorted by priority." } }]
        }
      },
      dbLink("Risks & Issues — filter: Status = Open or In Progress, group by Priority"),

      { object: "block", type: "divider", divider: {} },

      // ── Section 5: Recently Resolved ──
      {
        object: "block",
        type: "heading_2",
        heading_2: {
          rich_text: [{ type: "text", text: { content: "✅  Recently Resolved" } }]
        }
      },
      {
        object: "block",
        type: "paragraph",
        paragraph: {
          rich_text: [{ type: "text", text: { content: "Risks and issues closed recently. Good for status report evidence." } }]
        }
      },
      dbLink("Risks & Issues — filter: Status = Resolved, sort Resolved Date descending"),

      { object: "block", type: "divider", divider: {} },

      // ── Filter Notes ──
      {
        object: "block",
        type: "callout",
        callout: {
          icon: { type: "emoji", emoji: "💡" },
          rich_text: [
            {
              type: "text",
              text: {
                content: "How to use this dashboard:\n" +
                  "• Each database view above is a linked view — click it and apply filters in Notion.\n" +
                  "• Critical Path: filter Critical = ✓ AND % Complete < 100\n" +
                  "• Slipping Tasks: filter where Finish > Baseline Finish AND % Complete < 100\n" +
                  "• Milestones: filter Milestone = ✓, sort by Finish ascending\n" +
                  "• Open Risks: filter Status = Open or In Progress, group by Priority\n" +
                  "• Resolved: filter Status = Resolved, sort by Resolved Date descending"
              }
            }
          ]
        }
      }
    ]
  });

  console.log(`✅ Dashboard page created: ${page.id}`);
  console.log(`   Open it at: ${page.url}`);
  return page;
}

// ─── Main ────────────────────────────────────────────────────────────────────

(async () => {
  try {
    // Risks DB was already created — use its ID directly
    const risksDbId = process.env.NOTION_RISKS_DB_ID || await createRisksDatabase();
    await createDashboardPage(risksDbId);

    // Save IDs to env hint
    console.log("\n─────────────────────────────────────────");
    console.log("Add this to your .env:");
    console.log(`NOTION_RISKS_DB_ID=${risksDbId}`);
    console.log("─────────────────────────────────────────");
    console.log("\nDone! Open Notion and find 'Ametek Project Status — Management View' under the Ametek SAP S4 Implementation Project page.");
  } catch (err) {
    console.error("Setup failed:", err?.body ?? err);
    process.exit(1);
  }
})();

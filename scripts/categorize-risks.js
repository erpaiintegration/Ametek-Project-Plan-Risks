/* eslint-disable no-console */
/**
 * categorize-risks.js
 * Reads every record in the Risks & Issues DB, auto-assigns a detailed category
 * based on keyword matching against the record's title + any description text,
 * then writes the category back to Notion as a select property.
 *
 * Category property tried in order: "Risk Type", "Category", "Family"
 * Run:  node scripts/categorize-risks.js
 * Dry:  node scripts/categorize-risks.js --dry-run
 */
require("dotenv/config");
const { Client } = require("@notionhq/client");

const notion = new Client({ auth: process.env.NOTION_API_KEY });
const RISKS_DB_ID =
  process.env.NOTION_RISKS_DB_ID || "357ae9be-8a60-8190-b720-c130c7104cf1";
const DRY_RUN = process.argv.includes("--dry-run");

// ── Category rules (first match wins) ─────────────────────────────────────────
// Each rule explains WHY the category applies so the logic is auditable.
const RULES = [
  {
    category: "Data Migration",
    description:
      "Covers all data movement, conversion, cleansing, extract/load activities and legacy data readiness.",
    pattern:
      /data.?migr|migrat|extract|transform|cleansing|conversion|legacy.?data|data.?load|data.?quality|cutover.?data|ds4|master.?data|open.?item|balance.*carry|carry.*forward|bdc|lsmw|data.?prep/i,
  },
  {
    category: "Testing & UAT",
    description:
      "Covers all test cycles, UAT execution, defect resolution, regression, and sign-off activities.",
    pattern:
      /\btest\b|uat|user.?accept|validat|defect|error.?mitigat|quality.?assur|qa\b|regression|scenario|test.?case|test.?script|sit\b|unit.?test|integration.?test/i,
  },
  {
    category: "Integration & Interfaces",
    description:
      "Covers system-to-system connections, middleware, API, IDocs, EDI, and third-party connectors.",
    pattern:
      /integrat|interface|middleware|api\b|idoc|rfc\b|edi\b|connect|third.?party.?system|ext.?system|bi.?connector|boomi|mulesoft|sftp/i,
  },
  {
    category: "Go-Live & Cutover",
    description:
      "Covers production go-live execution, cutover planning, hypercare, rollback, and readiness gates.",
    pattern:
      /go.?live|cutover|hypercare|production.?start|readiness|deployment.?window|rollback|cut.?over|black.?out|freeze.?period|go live/i,
  },
  {
    category: "Training & Change Management",
    description:
      "Covers end-user training delivery, change management, communications, and user adoption.",
    pattern:
      /train|end.?user|user.?adopt|change.?manage|communic|awareness|e.?learn|learning|instructor|ocm\b|stakeholder.*aware/i,
  },
  {
    category: "Security & Access Control",
    description:
      "Covers role design, authorisations, SoD conflicts, Basis configuration, and access provisioning.",
    pattern:
      /security|access.?role|authoris|authoriz|permission|basis\b|role.?based|segreg|sod\b|grc\b|firefighter|super.?user|provisioning/i,
  },
  {
    category: "Reporting & Analytics",
    description:
      "Covers SAP reports, Fiori analytics, BW/BI content, dashboards, and output management.",
    pattern:
      /report|dashboard|analytic|bi\b|bw\b|fiori.?report|output|smartform|form.?template|correspondence|statement|sap.?analytic/i,
  },
  {
    category: "Configuration & Development",
    description:
      "Covers SAP configuration, custom development, ABAP enhancements, Fiori apps, and workflow.",
    pattern:
      /config|customis|customiz|develop|abap|fiori\b|workflow|enhancement|user.?exit|badi\b|gap|fit.?gap|bte\b|substitut|validation.?rule/i,
  },
  {
    category: "Infrastructure & Environment",
    description:
      "Covers system landscapes, transport management, performance, hardware, network, and cloud.",
    pattern:
      /infrastructure|environment|landscape|transport|performance|hardware|network|cloud|azure|aws|hosting|sizing|downtime|system.?copy|refresh/i,
  },
  {
    category: "Resource & Capacity",
    description:
      "Covers headcount availability, contractor gaps, skill shortages, and workload capacity.",
    pattern:
      /resource|staffing|headcount|capacit|availab|skill.?gap|contractor|consultant.?availab|team.?capacit|bandwidth/i,
  },
  {
    category: "PMO & Governance",
    description:
      "Covers steering committee decisions, scope changes, budget, approvals, and project governance.",
    pattern:
      /govern|steering|sign.?off|approval|pmo\b|budget|scope.?change|change.?request|escalat|decision|risk.?register|exec/i,
  },
  {
    category: "Schedule & Milestone Risk",
    description:
      "Covers timeline slippage, milestone delays, critical path pressure, and dependency cascade.",
    pattern:
      /schedul|timeline|delay|slip|milestone|critical.?path|depend|predecessor|successor|date.?change|replan/i,
  },
  {
    category: "Vendor & Third Party",
    description:
      "Covers SI partner delivery, software vendor support, licence issues, and external dependencies.",
    pattern:
      /vendor|third.?party|si\b|system.?integrat|partner|sap.?support|licence|license|contract|outsourc/i,
  },
];

function categorize(name, notes) {
  const text = `${name} ${notes || ""}`;
  for (const rule of RULES) {
    if (rule.pattern.test(text)) {
      return rule;
    }
  }
  return { category: "General Risk", description: "No specific keyword match — manual review recommended." };
}

function text(parts) {
  return (parts || []).map((p) => p.plain_text).join("");
}

async function fetchAll(dbId) {
  const out = [];
  let cursor;
  do {
    const r = await notion.databases.query({
      database_id: dbId,
      start_cursor: cursor,
      page_size: 100,
    });
    out.push(...r.results.filter((x) => x.object === "page"));
    cursor = r.has_more ? r.next_cursor : undefined;
  } while (cursor);
  return out;
}

async function main() {
  // ── 1. Inspect DB schema to find the right category property ───────────────
  console.log("Inspecting Risks & Issues database schema…");
  const db = await notion.databases.retrieve({ database_id: RISKS_DB_ID });
  const props = db.properties || {};

  console.log(
    "Available properties:\n" +
      Object.entries(props)
        .map(([n, p]) => `  ${n} (${p.type})`)
        .join("\n")
  );

  // Priority of category field candidates: select type, matching these names
  const candidates = ["Risk Type", "Category", "Family", "Risk Category", "Type Category"];
  let catProp = null;
  for (const c of candidates) {
    if (props[c]?.type === "select") {
      catProp = c;
      break;
    }
  }

  // If none found as select, pick first non-Type select field
  if (!catProp) {
    for (const [name, prop] of Object.entries(props)) {
      if (prop.type === "select" && !["Type", "Status", "Severity", "Priority", "Impact"].includes(name)) {
        catProp = name;
        break;
      }
    }
  }

  if (!catProp) {
    console.log('\n⚙  No category select property found — creating "Category" in the database…');
    await notion.databases.update({
      database_id: RISKS_DB_ID,
      properties: {
        Category: { select: {} },
      },
    });
    catProp = "Category";
    console.log('  ✅ "Category" property created.\n');
  }

  console.log(`\nUsing property "${catProp}" for category assignment.\n`);

  // ── 2. Fetch all risk/issue records ───────────────────────────────────────
  const pages = await fetchAll(RISKS_DB_ID);
  console.log(`Found ${pages.length} risk/issue records.\n`);

  const updates = [];

  for (const page of pages) {
    const p = page.properties || {};

    // Extract title
    const titleProp = Object.values(p).find((x) => x?.type === "title");
    const name = titleProp ? text(titleProp.title || []) : "(Untitled)";

    // Extract any description/notes text
    const notesProp = Object.values(p).find(
      (x) =>
        x?.type === "rich_text" &&
        ["Description", "Notes", "Details", "Summary", "Comments"].includes(
          Object.keys(p).find((k) => p[k] === x)
        )
    );
    const notes = notesProp ? text(notesProp.rich_text || []) : "";

    // Current category value
    const currentCat = p[catProp]?.select?.name || "(none)";

    // Categorize
    const match = categorize(name, notes);

    console.log(
      `  📋 "${name}"\n` +
        `     Was: ${currentCat}\n` +
        `     Now: ${match.category}\n` +
        `     Why: ${match.description}\n`
    );

    updates.push({ id: page.id, name, category: match.category, catProp });
  }

  if (DRY_RUN) {
    console.log("\n⚠  DRY RUN — no changes written to Notion.");
    console.log("Run without --dry-run to apply.");
    return;
  }

  // ── 3. Write categories back to Notion ────────────────────────────────────
  console.log("\nWriting categories to Notion…");
  let updated = 0;
  for (const u of updates) {
    try {
      await notion.pages.update({
        page_id: u.id,
        properties: {
          [u.catProp]: { select: { name: u.category } },
        },
      });
      updated++;
      console.log(`  ✅ ${u.name} → ${u.category}`);
    } catch (e) {
      console.error(`  ❌ Failed to update "${u.name}": ${e.message}`);
    }
  }

  console.log(`\nDone. ${updated}/${updates.length} records updated.`);
  console.log('Now run: npm run dashboard:html  to regenerate the dashboard.');
}

main().catch((e) => {
  console.error(e);
  process.exit(1);
});

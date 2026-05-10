/* eslint-disable no-console */
const fs = require("node:fs");
const path = require("node:path");
const { spawnSync } = require("node:child_process");

const ROOT = process.cwd();
const SHAREPOINT_DIR = path.join(ROOT, "From Ametek SharePoint");
const STAGING_DIR = path.join(ROOT, "imports", "staging");
const RAW_OUT = path.join(STAGING_DIR, "mpp-direct-raw.json");
const NORM_OUT = path.join(STAGING_DIR, "scheduling-tower-data.json");
const PS_SCRIPT = path.join(ROOT, "scripts", "export-mpp-json.ps1");

function parseArgs() {
  const out = {};
  for (let i = 2; i < process.argv.length; i += 1) {
    const token = process.argv[i];
    if (!token.startsWith("--")) continue;
    const key = token.slice(2);
    const val = process.argv[i + 1] && !process.argv[i + 1].startsWith("--") ? process.argv[i + 1] : true;
    out[key] = val;
    if (val !== true) i += 1;
  }
  return out;
}

function ensureDir(dir) {
  fs.mkdirSync(dir, { recursive: true });
}

function pickLatestMpp(dir) {
  const files = fs
    .readdirSync(dir)
    .filter((f) => f.toLowerCase().endsWith(".mpp"))
    .map((f) => {
      const full = path.join(dir, f);
      return { full, name: f, mtimeMs: fs.statSync(full).mtimeMs };
    })
    .sort((a, b) => b.mtimeMs - a.mtimeMs);

  if (!files.length) {
    throw new Error(`No .mpp files found in ${dir}`);
  }

  return files[0].full;
}

function daysDiff(a, b) {
  if (!a || !b) return null;
  const da = new Date(a);
  const db = new Date(b);
  if (Number.isNaN(da.getTime()) || Number.isNaN(db.getTime())) return null;
  return Math.round((da - db) / (24 * 60 * 60 * 1000));
}

function parseLinkCount(value) {
  if (!value) return 0;
  const nums = String(value).match(/\d+/g);
  return nums ? nums.length : 0;
}

function normalize(raw) {
  const now = new Date();
  const in14 = new Date(now.getTime() + 14 * 24 * 60 * 60 * 1000);

  const tasks = (raw.tasks || [])
    .filter((t) => t && t.uid != null && t.name)
    .map((t) => {
      const baselineFinish = t.baselineFinish || null;
      const slipDays = daysDiff(t.finish, baselineFinish);
      const isDone = Number(t.percentComplete || 0) >= 100;
      const isOpen = !isDone;
      const isSlipped = isOpen && slipDays != null && slipDays > 0;
      const isOverdueStart = isOpen && t.start ? new Date(t.start) < now : false;
      const isDue14 = isOpen && t.finish
        ? new Date(t.finish) >= now && new Date(t.finish) <= in14
        : false;

      return {
        uid: t.uid,
        id: t.id || null,
        name: t.name,
        wbs: t.wbs || t.outlineNumber || null,
        outlineNumber: t.outlineNumber || null,
        outlineLevel: Number.isFinite(Number(t.outlineLevel)) ? Number(t.outlineLevel) : null,
        summary: !!t.summary,
        milestone: !!t.milestone,
        critical: !!t.critical,
        percentComplete: Number(t.percentComplete || 0),
        start: t.start || null,
        finish: t.finish || null,
        baselineFinish,
        resourceNames: t.resourceNames || "",
        predecessors: t.predecessors || "",
        successors: t.successors || "",
        predCount: parseLinkCount(t.predecessors),
        sucCount: parseLinkCount(t.successors),
        slipDays,
        isOpen,
        isDone,
        isSlipped,
        isOverdueStart,
        isDue14
      };
    });

  const nonSummary = tasks.filter((t) => !t.summary);
  const open = nonSummary.filter((t) => t.isOpen);
  const slippedOpen = open.filter((t) => t.isSlipped);
  const overdueStarts = open.filter((t) => t.isOverdueStart);
  const due14 = open.filter((t) => t.isDue14);

  const slipRate = open.length ? (slippedOpen.length / open.length) * 100 : 0;

  const topSlipped = [...slippedOpen]
    .sort((a, b) => (b.slipDays || 0) - (a.slipDays || 0))
    .slice(0, 100);
  const topDue14 = [...due14]
    .sort((a, b) => (a.finish || "").localeCompare(b.finish || ""))
    .slice(0, 100);
  const topOverdueStarts = [...overdueStarts]
    .sort((a, b) => (a.start || "").localeCompare(b.start || ""))
    .slice(0, 100);

  return {
    generatedAt: new Date().toISOString(),
    source: {
      type: "mpp",
      file: raw.sourceFile,
      projectName: raw.projectName,
      exportedAt: raw.exportedAt
    },
    metrics: {
      totalTasks: nonSummary.length,
      openTasks: open.length,
      doneTasks: nonSummary.length - open.length,
      slippedOpen: slippedOpen.length,
      slipRatePct: Number(slipRate.toFixed(1)),
      due14: due14.length,
      overdueStarts: overdueStarts.length,
      criticalTasks: nonSummary.filter((t) => t.critical && t.isOpen).length,
      milestonesOpen: nonSummary.filter((t) => t.milestone && t.isOpen).length
    },
    lists: {
      topSlipped,
      topDue14,
      topOverdueStarts
    },
    tasks: nonSummary
  };
}

function runPowerShell(mppPath, outJson, options = {}) {
  const psArgs = [
    "-NoProfile",
    "-ExecutionPolicy",
    "Bypass",
    "-File",
    PS_SCRIPT,
    "-MppPath",
    mppPath,
    "-OutJson",
    outJson
  ];

  if (options.useActiveProject) {
    psArgs.push("-UseActiveProject");
  }

  const result = spawnSync(
    "powershell.exe",
    psArgs,
    { encoding: "utf8" }
  );

  if (result.status !== 0) {
    throw new Error(`PowerShell export failed (code ${result.status}):\n${result.stderr || result.stdout}`);
  }
}

function main() {
  const args = parseArgs();
  ensureDir(STAGING_DIR);

  const mppPath = args.file
    ? path.isAbsolute(args.file)
      ? args.file
      : path.join(ROOT, args.file)
    : pickLatestMpp(SHAREPOINT_DIR);
  const useActiveProject = args.active === true || args["use-active-project"] === true;

  console.log(`MPP source: ${mppPath}`);
  if (useActiveProject) {
    console.log("Using active MS Project window for extraction.");
  }
  runPowerShell(mppPath, RAW_OUT, { useActiveProject });

  const raw = JSON.parse(fs.readFileSync(RAW_OUT, "utf8"));
  const normalized = normalize(raw);

  fs.writeFileSync(NORM_OUT, JSON.stringify(normalized), "utf8");

  console.log("Direct .mpp ingest complete:");
  console.log(`  Project: ${normalized.source.projectName || "(unknown)"}`);
  console.log(`  Tasks: ${normalized.metrics.totalTasks}`);
  console.log(`  Open: ${normalized.metrics.openTasks}`);
  console.log(`  Slipped Open: ${normalized.metrics.slippedOpen}`);
  console.log(`  Slip Rate: ${normalized.metrics.slipRatePct}%`);
  console.log(`  Output: ${NORM_OUT}`);
}

main();

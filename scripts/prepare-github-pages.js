/* eslint-disable no-console */
require("dotenv/config");
const fs = require("fs");
const path = require("path");

const ROOT = path.join(__dirname, "..");
const DASHBOARD_FILE = path.join(ROOT, "dashboard.html");
const DASHBOARD_DATA_FILE = path.join(ROOT, "dashboard-data.json");
const DOCS_DIR = path.join(ROOT, "docs");
const INDEX_FILE = path.join(DOCS_DIR, "index.html");
const DOCS_DATA_FILE = path.join(DOCS_DIR, "dashboard-data.json");
const NOJEKYLL_FILE = path.join(DOCS_DIR, ".nojekyll");

function ensureDir(dirPath) {
  fs.mkdirSync(dirPath, { recursive: true });
}

function main() {
  if (!fs.existsSync(DASHBOARD_FILE)) {
    throw new Error(
      `dashboard.html not found at ${DASHBOARD_FILE}. Run the dashboard generator first.`
    );
  }

  if (!fs.existsSync(DASHBOARD_DATA_FILE)) {
    throw new Error(
      `dashboard-data.json not found at ${DASHBOARD_DATA_FILE}. Run the dashboard generator first.`
    );
  }

  ensureDir(DOCS_DIR);
  fs.copyFileSync(DASHBOARD_FILE, INDEX_FILE);
  fs.copyFileSync(DASHBOARD_DATA_FILE, DOCS_DATA_FILE);

  if (!fs.existsSync(NOJEKYLL_FILE)) {
    fs.writeFileSync(NOJEKYLL_FILE, "", "utf8");
  }

  console.log(`GitHub Pages site prepared:`);
  console.log(`  Source: ${DASHBOARD_FILE}`);
  console.log(`  Publish: ${INDEX_FILE}`);
  console.log(`  Data: ${DOCS_DATA_FILE}`);
  console.log(`  Marker: ${NOJEKYLL_FILE}`);
}

main();

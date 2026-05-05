/* eslint-disable no-console */
const { execSync } = require("child_process");

function run(cmd) {
  console.log(`\n> ${cmd}`);
  execSync(cmd, { stdio: "inherit" });
}

function safe(cmd) {
  try {
    return execSync(cmd, { stdio: ["ignore", "pipe", "pipe"] }).toString().trim();
  } catch {
    return "";
  }
}

function main() {
  run("npm run dashboard:pages");

  run("git add docs/index.html");

  const changed = safe("git diff --cached --name-only");
  if (!changed) {
    console.log("\nNo dashboard changes detected. Nothing to commit.");
    return;
  }

  const stamp = new Date().toISOString().replace("T", " ").slice(0, 16);
  run(`git commit -m "chore: dashboard snapshot ${stamp}"`);
  run("git push");

  console.log("\nPublished dashboard snapshot to GitHub.");
}

main();

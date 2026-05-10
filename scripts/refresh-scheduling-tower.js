/* eslint-disable no-console */
const { spawnSync } = require("node:child_process");

function runNode(script, args = []) {
  const result = spawnSync(process.execPath, [script, ...args], {
    stdio: "inherit",
    encoding: "utf8"
  });
  if (result.status !== 0) {
    throw new Error(`Failed: ${script}`);
  }
}

function main() {
  const passthrough = process.argv.slice(2);
  runNode("scripts/ingest-mpp-direct.js", passthrough);
  runNode("scripts/generate-scheduling-tower-standalone.js");
  console.log("Scheduling Tower refresh completed.");
}

main();

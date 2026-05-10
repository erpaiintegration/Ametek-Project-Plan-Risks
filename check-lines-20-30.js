const fs = require('fs');
const html = fs.readFileSync('dashboard.html', 'utf8');
const scriptStart = html.indexOf('<script>') + 8;
const scriptEnd = html.indexOf('</script>');
const fullScript = html.slice(scriptStart, scriptEnd);
const lines = fullScript.split('\n');

console.log('Lines 20-30 (full, no truncation):');
for (let i = 19; i < 30 && i < lines.length; i++) {
  const line = lines[i];
  const len = line.length;
  console.log(`${i + 1} (len=${len}): ${line.substring(0, 150)}${len > 150 ? '...' : ''}`);
}

// Test if there's something wonky with the newlines
console.log('\n--- Checking newline characters ---');
for (let i = 19; i < 25; i++) {
  const line = lines[i];
  console.log(`Line ${i + 1}: starts with 0x${lines[i].charCodeAt(0)?.toString(16).padStart(2, '0')}, ends with 0x${lines[i].charCodeAt(lines[i].length - 1)?.toString(16).padStart(2, '0')}`);
}

// Check if we can manually test what happens
console.log('\n--- Manual test ---');
const manualCode = `
const DATA = {"a":1};
let activeFilter = null;
const fmt = d => d ? "test" : "—";
const PERF = { initialTaskRows: 180, maxTaskRows: 320 };
let atRiskTasksCache = null;

async function loadTaskData() {
  console.log('test');
}
`;

try {
  new Function(manualCode);
  console.log('✓ Manual code is valid');
} catch (e) {
  console.log('✗ Manual code error:', e.message);
}

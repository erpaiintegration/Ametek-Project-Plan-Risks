const fs = require('fs');

const html = fs.readFileSync('./dashboard.html', 'utf8');
const lines = html.split('\n');

// Find the script tag
let scriptStart = -1;
for (let i = 0; i < lines.length; i++) {
  if (lines[i].includes('<script>')) {
    scriptStart = i;
    break;
  }
}

const scriptLines = lines.slice(scriptStart + 1);

console.log('Testing which line between 2-20 causes the issue...\n');

// Test lines 1-2, 1-3, 1-4, etc. until we find the problem
for (let i = 2; i <= 20; i++) {
  const testCode = scriptLines.slice(0, i).join('\n');
  try {
    new Function(testCode);
    console.log(`✓ Lines 1-${i} parse successfully`);
  } catch (e) {
    console.log(`✗ Lines 1-${i} ERROR: ${e.message}`);
    console.log(`  Problem line (${i}): ${scriptLines[i-1].substring(0, 100)}`);
    break;
  }
}

// Now let's look at the specific content of lines 2-20
console.log(`\n--- Content of lines 2-20 ---`);
for (let i = 1; i < Math.min(20, scriptLines.length); i++) {
  const line = scriptLines[i].length > 120 ? scriptLines[i].substring(0, 120) + '...' : scriptLines[i];
  console.log(`${i + 1}: ${line}`);
}

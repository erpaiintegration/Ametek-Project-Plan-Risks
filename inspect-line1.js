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
const line1 = scriptLines[0]; // This should be the "const DATA = ..." line

console.log(`Line 1 length: ${line1.length} characters`);
console.log(`\nFirst 100 characters:\n${line1.substring(0, 100)}`);
console.log(`\nLast 200 characters:\n${line1.substring(Math.max(0, line1.length - 200))}`);

// Try to parse just this line
console.log(`\n--- Testing line 1 alone ---`);
try {
  new Function(line1);
  console.log('✓ Line 1 parses successfully!');
} catch (e) {
  console.log(`✗ Error: ${e.message}`);
}

// Try parsing line 1 + line 2
console.log(`\n--- Testing line 1 + line 2 ---`);
try {
  new Function(line1 + '\n' + scriptLines[1]);
  console.log('✓ Lines 1-2 parse successfully!');
} catch (e) {
  console.log(`✗ Error: ${e.message}`);
}

// Let's check if line 1 is a valid variable assignment
console.log(`\n--- Checking DATA assignment structure ---`);
const dataMatch = line1.match(/^const DATA = (.+);$/);
if (dataMatch) {
  console.log('✓ Line matches: const DATA = {...};');
  const jsonPart = dataMatch[1];
  console.log(`JSON content length: ${jsonPart.length}`);
  
  // Try to parse just the JSON
  try {
    JSON.parse(jsonPart);
    console.log('✓ JSON parses successfully!');
  } catch (e) {
    console.log(`✗ JSON parse error: ${e.message}`);
  }
} else {
  console.log('✗ Line 1 does NOT match expected format!');
  // Check what it actually looks like
  console.log(`First 50 chars: ${line1.substring(0, 50)}`);
  console.log(`Last 50 chars: ${line1.substring(line1.length - 50)}`);
}

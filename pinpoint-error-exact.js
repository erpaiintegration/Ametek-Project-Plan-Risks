const fs = require('fs');

// Load dashboard.html and extract the actual script
const html = fs.readFileSync('./dashboard.html', 'utf8');
const scriptStart = html.indexOf('<script>') + '<script>'.length;
const scriptEnd = html.indexOf('</script>');
const rawScript = html.substring(scriptStart, scriptEnd);

const lines = rawScript.split('\n');
console.log(`Total lines: ${lines.length}`);

// Binary search to find exact error line
let low = 0;
let high = lines.length;

while (low < high) {
  const mid = Math.floor((low + high) / 2);
  const testCode = lines.slice(0, mid + 1).join('\n');
  
  try {
    new Function(testCode);
    low = mid + 1;
  } catch (err) {
    high = mid;
  }
}

console.log(`\nError first appears at line ${low + 1}`);
console.log(`Lines 1-${low} parse successfully`);
console.log(`Lines 1-${low + 1} fail to parse\n`);

// Show lines around the error
const startLine = Math.max(0, low - 3);
const endLine = Math.min(lines.length, low + 5);

console.log(`Context (lines ${startLine + 1} to ${endLine}):`);
for (let i = startLine; i < endLine; i++) {
  const prefix = i === low ? '>>> ' : '    ';
  const line = lines[i];
  const display = line.length > 100 ? line.substring(0, 100) + '...' : line;
  console.log(`${prefix}${i + 1}: ${display}`);
}

// Try to get the exact error from line low alone
console.log(`\n--- Line ${low + 1} alone (without context) ---`);
try {
  new Function(lines[low]);
  console.log('✓ Line parses when isolated');
} catch (e) {
  console.log(`✗ Line fails when isolated: ${e.message}`);
}

// Try lines 1-20
console.log(`\n--- Lines 1-20 combined ---`);
try {
  new Function(lines.slice(0, 20).join('\n'));
  console.log('✓ Lines 1-20 parse successfully');
} catch (e) {
  console.log(`✗ ERROR: ${e.message}`);
}

// Try lines 1-21
console.log(`\n--- Lines 1-21 combined ---`);
try {
  new Function(lines.slice(0, 21).join('\n'));
  console.log('✓ Lines 1-21 parse successfully');
} catch (e) {
  console.log(`✗ ERROR: ${e.message}`);
}

// Try lines 1-22
console.log(`\n--- Lines 1-22 combined ---`);
try {
  new Function(lines.slice(0, 22).join('\n'));
  console.log('✓ Lines 1-22 parse successfully');
} catch (e) {
  console.log(`✗ ERROR: ${e.message}`);
}

// Show exact content of lines 20-22
console.log(`\n--- Exact line content (lines 20-22) ---`);
for (let i = 19; i < 22; i++) {
  if (i < lines.length) {
    const line = lines[i];
    console.log(`Line ${i + 1} (${line.length} chars): ${JSON.stringify(line)}`);
  }
}

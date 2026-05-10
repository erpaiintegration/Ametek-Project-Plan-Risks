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

// Find the closing script tag
let scriptEnd = -1;
for (let i = scriptStart; i < lines.length; i++) {
  if (lines[i].includes('</script>')) {
    scriptEnd = i;
    break;
  }
}

// Extract script lines
const scriptLines = lines.slice(scriptStart + 1, scriptEnd);

console.log(`Testing ${scriptLines.length} lines of script content...\n`);

// Binary search for the error
let low = 0;
let high = scriptLines.length;

while (low < high) {
  const mid = Math.floor((low + high) / 2);
  const testCode = scriptLines.slice(0, mid + 1).join('\n');
  
  try {
    new Function(testCode);
    // Code parsed successfully, error is further down
    low = mid + 1;
  } catch (e) {
    // Code failed to parse, error is in this chunk
    high = mid;
  }
}

console.log(`✓ First ${low} lines parse successfully`);
console.log(`✗ Error occurs at line ${low + 1} of script (file line ${scriptStart + 1 + low + 1})`);

// Show context around the error
console.log(`\n--- Content around error line ---`);
for (let i = Math.max(0, low - 3); i <= Math.min(scriptLines.length - 1, low + 2); i++) {
  const prefix = i === low ? '>>> ' : '    ';
  const line = scriptLines[i].length > 150 ? scriptLines[i].substring(0, 150) + '...' : scriptLines[i];
  console.log(`${prefix}${i + 1}: ${line}`);
}

// Try to parse just the problem line with context
console.log(`\n--- Trying to isolate the error ---`);
if (low > 0) {
  try {
    const testCode = scriptLines.slice(0, low).join('\n');
    new Function(testCode);
    console.log(`✓ Lines 1-${low} parse successfully`);
  } catch (e) {
    console.log(`✗ Error in lines 1-${low}: ${e.message}`);
  }
}

try {
  const testCode = scriptLines.slice(0, low + 1).join('\n');
  new Function(testCode);
  console.log(`✓ Lines 1-${low + 1} parse successfully (unexpected!)`);
} catch (e) {
  console.log(`✗ Error in lines 1-${low + 1}: ${e.message}`);
  
  // Show the exact content of the problem line
  console.log(`\nLine ${low + 1} content (length ${scriptLines[low].length}):`);
  const line = scriptLines[low];
  console.log(line.substring(0, 200));
  if (line.length > 200) console.log('...[truncated]...');
  if (line.length > 200) console.log(line.substring(Math.max(0, line.length - 200)));
}

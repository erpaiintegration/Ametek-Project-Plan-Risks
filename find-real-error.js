const fs = require('fs');

// Load dashboard.html
const html = fs.readFileSync('./dashboard.html', 'utf8');
const scriptOpenTag = html.indexOf('<script>');
const scriptCloseTag = html.indexOf('</script>');
const scriptContent = html.substring(scriptOpenTag + '<script>'.length, scriptCloseTag);

// Split by newlines
const lines = scriptContent.split('\n');

console.log(`Total lines: ${lines.length}`);
console.log(`\n=== First 25 lines (with cumulative character counts) ===`);

let charCount = 0;
for (let i = 0; i < Math.min(25, lines.length); i++) {
  const line = lines[i];
  charCount += line.length + 1; // +1 for newline
  const display = line.length > 80 ? line.substring(0, 80) + '... (' + line.length + ' chars)' : line;
  const marker = line.includes('async function loadTaskData') ? '👉 ' : '   ';
  console.log(`${marker}Line ${i + 1}: ${display}`);
  if (line.includes('async function loadTaskData')) {
    console.log(`\n✓ Found async function on line ${i + 1}!`);
    break;
  }
}

// Now binary search for the actual error
console.log(`\n=== Binary search for parse error ===`);
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

console.log(`First ${low} lines parse successfully`);
console.log(`First ${low + 1} lines fail to parse\n`);

// Show context around error
const errorLine = low;
const start = Math.max(0, errorLine - 3);
const end = Math.min(lines.length, errorLine + 5);

console.log(`Context (lines ${start + 1} to ${end}):`);
for (let i = start; i < end; i++) {
  const line = lines[i];
  const prefix = i === errorLine ? '>>> ' : '    ';
  const display = line.length > 100 ? line.substring(0, 100) + '...' : line;
  console.log(`${prefix}Line ${i + 1}: ${display}`);
}

// Try to get the error
console.log(`\n--- Error test ---`);
const testCode = lines.slice(0, errorLine + 1).join('\n');
try {
  new Function(testCode);
} catch (e) {
  console.log(`Error: ${e.message}`);
}

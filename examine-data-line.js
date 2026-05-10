const fs = require('fs');

// Load dashboard.html
const html = fs.readFileSync('./dashboard.html', 'utf8');
const scriptOpenTag = html.indexOf('<script>');
const scriptCloseTag = html.indexOf('</script>');
const scriptContent = html.substring(scriptOpenTag + '<script>'.length, scriptCloseTag);

// Split by newlines
const lines = scriptContent.split('\n');

// Examine line 2 (the DATA line)
const dataLine = lines[1]; // Line 2 is index 1

console.log(`Line 2 (DATA) length: ${dataLine.length}`);
console.log(`First 50 chars: ${dataLine.substring(0, 50)}`);
console.log(`Last 100 chars: ${dataLine.substring(dataLine.length - 100)}`);

// Check if it starts with "const DATA = {" and ends with "};"
console.log(`\nStarts with "const DATA = {": ${dataLine.substring(0, 15)}`);
console.log(`Ends with "};": ${dataLine.substring(dataLine.length - 5)}`);

// Test if this line alone parses
console.log(`\n--- Test line 2 alone ---`);
try {
  new Function(dataLine);
  console.log('✓ Line 2 alone parses!');
} catch (e) {
  console.log(`✗ Error: ${e.message}`);
}

// Test lines 1-2
console.log(`\n--- Test lines 1-2 combined ---`);
try {
  new Function(lines.slice(0, 2).join('\n'));
  console.log('✓ Lines 1-2 parse!');
} catch (e) {
  console.log(`✗ Error: ${e.message}`);
}

// Test lines 1-21
console.log(`\n--- Test lines 1-21 combined ---`);
try {
  new Function(lines.slice(0, 21).join('\n'));
  console.log('✓ Lines 1-21 parse!');
} catch (e) {
  console.log(`✗ Error: ${e.message}`);
}

// Now let me check the EXACT character at the end
console.log(`\n--- Last 20 char codes of line 2 ---`);
const lastChars = dataLine.substring(dataLine.length - 20);
for (let i = 0; i < lastChars.length; i++) {
  const ch = lastChars[i];
  console.log(`Pos ${dataLine.length - 20 + i}: '${ch}' (code ${ch.charCodeAt(0)})`);
}

// Check for any special characters
console.log(`\n--- Check for problematic characters in line 2 ---`);
const problematicChars = ['"', "'", '/', '\\', '\r', '\0', '\t'];
for (const ch of problematicChars) {
  const count = (dataLine.match(new RegExp(ch === '\\' ? '\\\\' : ch.replace(/[.*+?^${}()|[\]\\]/g, '\\$&'), 'g')) || []).length;
  if (count > 0) {
    console.log(`Found ${count} instances of '${ch}'`);
  }
}

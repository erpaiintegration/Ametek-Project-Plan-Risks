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

console.log(`Line 20: "${scriptLines[19]}"`);
console.log(`Line 20 length: ${scriptLines[19].length}`);
console.log(`Line 20 charCodes: ${[...scriptLines[19]].map(c => c.charCodeAt(0)).join(', ')}`);

console.log(`\nLine 21: "${scriptLines[20]}"`);
console.log(`Line 21 length: ${scriptLines[20].length}`);
console.log(`Line 21 first 50 chars: ${scriptLines[20].substring(0, 50)}`);
console.log(`Line 21 charCodes (first 50): ${[...scriptLines[20].substring(0, 50)].map(c => c.charCodeAt(0)).join(', ')}`);

console.log(`\nLine 22: "${scriptLines[21]}"`);

// Now test adding just line 21
console.log(`\n--- Test: Lines 1-20 + Line 21 ---`);
const code1to20 = scriptLines.slice(0, 20).join('\n');
try {
  new Function(code1to20);
  console.log('✓ Lines 1-20 parse');
} catch (e) {
  console.log(`✗ Lines 1-20 error: ${e.message}`);
}

const code1to21 = scriptLines.slice(0, 21).join('\n');
try {
  new Function(code1to21);
  console.log('✓ Lines 1-21 parse');
} catch (e) {
  console.log(`✗ Lines 1-21 error: ${e.message}`);
}

// Let's also check if there's something funky with line 20's ending
// Maybe there's a missing semicolon somewhere
console.log(`\n--- Checking line endings ---`);
for (let i = 18; i < 23 && i < scriptLines.length; i++) {
  const line = scriptLines[i];
  const lastChar = line[line.length - 1] ? line.charCodeAt(line.length - 1) : 'N/A';
  const lastThree = line.substring(Math.max(0, line.length - 10));
  console.log(`Line ${i+1}: ends with '${lastThree}' (charCode ${lastChar})`);
}

// Check if line 21 actually starts with "async"
console.log(`\n--- Checking line 21 syntax ---`);
const line21 = scriptLines[20];
console.log(`Starts with 'async': ${line21.startsWith('async')}`);
console.log(`Full line: ${line21}`);

// Try wrapping it in additional context
console.log(`\n--- Test: Add semicolon after line 20 ---`);
const codeWithSemi = scriptLines.slice(0, 20).join('\n') + ';\n' + scriptLines[20] + '\n' + scriptLines[21];
try {
  new Function(codeWithSemi);
  console.log('✓ Lines 1-22 WITH added semicolon parse');
} catch (e) {
  console.log(`✗ Error: ${e.message}`);
}

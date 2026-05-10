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
const line1 = scriptLines[0];

console.log(`Line 1 total length: ${line1.length}`);

// Get the last 500 characters
const last500 = line1.substring(Math.max(0, line1.length - 500));
console.log(`\n--- Last 500 characters ---`);
console.log(last500);

// Get character codes of the last 50 characters
const last50 = line1.substring(Math.max(0, line1.length - 50));
console.log(`\n--- Last 50 characters (with char codes) ---`);
for (let i = 0; i < last50.length; i++) {
  const char = last50[i];
  const code = char.charCodeAt(0);
  const display = char === '\n' ? '\\n' : char === '\r' ? '\\r' : char === '\t' ? '\\t' : char;
  console.log(`[${i}] '${display}' (charCode ${code})`);
}

// Check if line 1 ends with "};
console.log(`\n--- Line 1 ending analysis ---`);
console.log(`Ends with '};': ${line1.endsWith('};')}`);
console.log(`Ends with '}': ${line1.endsWith('}')}`);
console.log(`Ends with '];': ${line1.endsWith('];')}`);
console.log(`Ends with ']}': ${line1.endsWith(']}')}`);

// Count braces and brackets
let braceCount = 0;
let bracketCount = 0;
for (let i = 0; i < line1.length; i++) {
  if (line1[i] === '{') braceCount++;
  else if (line1[i] === '}') braceCount--;
  else if (line1[i] === '[') bracketCount++;
  else if (line1[i] === ']') bracketCount--;
}

console.log(`\n--- Balance check ---`);
console.log(`Unclosed braces: ${braceCount}`);
console.log(`Unclosed brackets: ${bracketCount}`);

// Also check if the line is valid as a complete statement
console.log(`\n--- Validity check ---`);
try {
  eval(line1 + ';');
  console.log('✓ Line 1 is valid JavaScript when evaluated');
} catch (e) {
  console.log(`✗ Line 1 evaluation error: ${e.message}`);
}

// Check if it's valid within a function context
try {
  new Function('return ' + line1);
  console.log('✓ Line 1 can be used in return statement');
} catch (e) {
  console.log(`✗ Error: ${e.message}`);
}

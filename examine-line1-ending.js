const fs = require('fs');

// Load dashboard.html and extract the actual script
const html = fs.readFileSync('./dashboard.html', 'utf8');
const scriptStart = html.indexOf('<script>') + '<script>'.length;
const scriptEnd = html.indexOf('</script>');
const rawScript = html.substring(scriptStart, scriptEnd);

const lines = rawScript.split('\n');

// Get line 1
const line1 = lines[0];
console.log(`Line 1 length: ${line1.length}`);
console.log(`Line 1 starts with: ${line1.substring(0, 50)}`);
console.log(`Line 1 ends with: ${line1.substring(line1.length - 50)}`);

// Count braces/brackets in line 1
let openBraces = 0, closeBraces = 0;
let openBrackets = 0, closeBrackets = 0;
let openParens = 0, closeParens = 0;

for (let i = 0; i < line1.length; i++) {
  const ch = line1[i];
  if (ch === '{') openBraces++;
  else if (ch === '}') closeBraces++;
  else if (ch === '[') openBrackets++;
  else if (ch === ']') closeBrackets++;
  else if (ch === '(') openParens++;
  else if (ch === ')') closeParens++;
}

console.log(`\nBrace balance in line 1:`);
console.log(`  Braces: ${openBraces} open, ${closeBraces} close (net: ${openBraces - closeBraces})`);
console.log(`  Brackets: ${openBrackets} open, ${closeBrackets} close (net: ${openBrackets - closeBrackets})`);
console.log(`  Parens: ${openParens} open, ${closeParens} close (net: ${openParens - closeParens})`);

// The problem might be an unclosed paren. Let me check line 1 character by character near the end
console.log(`\n--- Last 200 characters of line 1 (with positions) ---`);
const lastChars = line1.substring(Math.max(0, line1.length - 200));
for (let i = 0; i < lastChars.length; i++) {
  const ch = lastChars[i];
  if (ch === '{' || ch === '}' || ch === '[' || ch === ']' || ch === '(' || ch === ')' || ch === ';' || ch === ',' || ch === ':' || ch === '"' || ch === "'") {
    console.log(`Pos ${line1.length - 200 + i}: '${ch}'`);
  }
}

// Try parsing just line 1
console.log(`\n--- Attempting to parse line 1 as code ---`);
try {
  new Function(line1);
  console.log('✓ Line 1 parses successfully!');
} catch (e) {
  console.log(`✗ ERROR: ${e.message}`);
}

// Maybe the issue is the way the template joins them? Let me test
console.log(`\n--- Testing with proper variable wrapper ---`);
try {
  new Function('{ ' + line1 + ' }');
  console.log('✓ Parses with wrapper');
} catch (e) {
  console.log(`✗ ERROR with wrapper: ${e.message}`);
}

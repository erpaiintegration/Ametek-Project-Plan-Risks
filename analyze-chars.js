const fs = require('fs');
const html = fs.readFileSync('dashboard.html', 'utf8');
const scriptStart = html.indexOf('<script>') + 8;
const scriptEnd = html.indexOf('</script>');
const script = html.slice(scriptStart, scriptEnd);
const lines = script.split('\n');

// Check for unbalanced parentheses across cumulative lines
console.log('=== Parenthesis balance check by line ===');
let openParens = 0, openBrackets = 0, openBraces = 0;
for (let i = 0; i <= 20; i++) {
  const line = lines[i];
  const prevParens = openParens;
  for (const char of line) {
    if (char === '(') openParens++;
    if (char === ')') openParens--;
    if (char === '[') openBrackets++;
    if (char === ']') openBrackets--;
    if (char === '{') openBraces++;
    if (char === '}') openBraces--;
  }
  if (i <= 16 || openParens !== prevParens) {
    const linePreview = line.substring(0, 50);
    console.log(`Line ${i + 1}: parens=${prevParens}→${openParens}, brackets=${openBrackets}, braces=${openBraces} | ${linePreview}...`);
  }
}

console.log(`\nFinal: parens=${openParens}, brackets=${openBrackets}, braces=${openBraces}`);
console.log(`\nLines 1-20 cumulative:`);
try {
  new Function(lines.slice(0, 21).join('\n'));
  console.log('VALID');
} catch (e) {
  console.log(`ERROR: ${e.message}`);
}

console.log(`\nLines 1-21 cumulative:`);
try {
  new Function(lines.slice(0, 22).join('\n'));
  console.log('VALID');
} catch (e) {
  console.log(`ERROR: ${e.message}`);
}

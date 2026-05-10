const fs = require('fs');
const html = fs.readFileSync('dashboard.html', 'utf8');
const scriptStart = html.indexOf('<script>') + 8;
const scriptEnd = html.indexOf('</script>');
const script = html.slice(scriptStart, scriptEnd);
const lines = script.split('\n');

const line2 = lines[1];
console.log('Line 2 (DATA line):');
console.log('  First 100 chars:', line2.substring(0, 100));
console.log('  Last 100 chars:', line2.substring(line2.length - 100));

// Check for balanced parens/braces in first 2 lines
let openParens = 0, openBraces = 0, inString = false, stringChar = '';
for (let charIdx = 0; charIdx < line2.length; charIdx++) {
  const char = line2[charIdx];
  
  if (!inString && (char === '"' || char === "'" || char === '`')) {
    inString = true;
    stringChar = char;
  } else if (inString && char === stringChar && line2[charIdx - 1] !== '\\') {
    inString = false;
  } else if (!inString) {
    if (char === '(') openParens++;
    if (char === ')') openParens--;
    if (char === '{') openBraces++;
    if (char === '}') openBraces--;
  }
}

console.log(`\nLine 2 balance:`);
console.log(`  Open parens: ${openParens}`);
console.log(`  Open braces: ${openBraces}`);

// Check all lines 1-21
console.log(`\n--- Cumulative line balance (lines 1-21) ---`);
let totalParens = 0, totalBraces = 0;
for (let lineIdx = 0; lineIdx < 21; lineIdx++) {
  const line = lines[lineIdx];
  let parens = 0, braces = 0;
  let inStr = false, strCh = '';
  
  for (let charIdx = 0; charIdx < line.length; charIdx++) {
    const char = line[charIdx];
    
    if (!inStr && (char === '"' || char === "'" || char === '`')) {
      inStr = true;
      strCh = char;
    } else if (inStr && char === strCh && line[charIdx - 1] !== '\\') {
      inStr = false;
    } else if (!inStr) {
      if (char === '(') parens++;
      if (char === ')') parens--;
      if (char === '{') braces++;
      if (char === '}') braces--;
    }
  }
  
  totalParens += parens;
  totalBraces += braces;
  if (lineIdx < 12 || lineIdx >= 18) {
    console.log(`Line ${lineIdx + 1}: p${parens > 0 ? '+' : ''}${parens}, b${braces > 0 ? '+' : ''}${braces} | Total: p${totalParens}, b${totalBraces}`);
  } else if (lineIdx === 12) {
    console.log(`... (lines 13-18 omitted)`);
  }
}

console.log(`\nFinal totals for lines 1-21:`);
console.log(`  Unclosed parens: ${totalParens}`);
console.log(`  Unclosed braces: ${totalBraces}`);

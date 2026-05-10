const fs = require('fs');
const html = fs.readFileSync('dashboard.html', 'utf8');
const scriptStart = html.indexOf('<script>') + 8;
const scriptEnd = html.indexOf('</script>');
const script = html.slice(scriptStart, scriptEnd);

// Extract the DATA part
const dataStart = script.indexOf('const DATA = ');
const endSemicolon = script.indexOf(';', dataStart);
const jsonStr = script.slice(dataStart + 13, endSemicolon);

const pos = 883;
console.log('Context around position 883:');
console.log('String segment (850-900):');
console.log(jsonStr.substring(850, 900));

console.log('\nCharacter codes (880-890):');
for (let i = 880; i < 890; i++) {
  const char = jsonStr[i];
  const code = jsonStr.charCodeAt(i);
  console.log(`  ${i}: '${char === '"' ? '\\"' : char}' (code ${code})`);
}

// Find all unescaped quotes
console.log('\n--- Checking quote balance in first 900 characters ---');
let inString = false, i = 0;
while (i < 900) {
  const char = jsonStr[i];
  const prevChar = jsonStr[i - 1];
  
  if (char === '"' && prevChar !== '\\') {
    inString = !inString;
    if (!inString) console.log(`  Quote closed at ${i}`);
    else console.log(`  Quote opened at ${i}`);
  }
  i++;
}

console.log(`\nAt position 900, inString state: ${inString}`);

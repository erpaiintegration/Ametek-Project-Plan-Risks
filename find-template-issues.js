// Find all literal newlines inside single/double-quoted strings in the generator template
const fs = require('fs');
const gen = fs.readFileSync('scripts/generate-dashboard-html.js', 'utf8');
const lines = gen.split('\n');

// Find the template literal boundaries (return `...` in generateHTML function)
let templateStart = -1, templateEnd = -1;
for (let i = 0; i < lines.length; i++) {
  if (templateStart === -1 && lines[i].match(/return\s+`/)) {
    templateStart = i;
  }
}

console.log('Template starts at generator line:', templateStart + 1);

// Look for patterns like: line ends with ' or " (suggesting open string continuing to next line)
// within the template
const issues = [];
let inTemplate = templateStart >= 0;

for (let i = templateStart; i < lines.length - 1; i++) {
  const line = lines[i];
  const nextLine = lines[i + 1];
  
  // Skip template literal interpolations ${...}
  // Look for a single quote that opens a string that doesn't close on the same line
  // Simple heuristic: count unescaped quotes
  let singleCount = 0, doubleCount = 0;
  let j = 0;
  while (j < line.length) {
    if (line[j] === '\\') { j += 2; continue; }
    if (line[j] === "'") singleCount++;
    if (line[j] === '"') doubleCount++;
    j++;
  }
  
  // Odd number of single quotes means there's an unclosed string
  if (singleCount % 2 === 1) {
    issues.push({
      genLine: i + 1,
      content: line.trimEnd(),
      next: nextLine.trimEnd(),
      type: 'single-quote'
    });
  }
  if (doubleCount % 2 === 1) {
    issues.push({
      genLine: i + 1,
      content: line.trimEnd(),
      next: nextLine.trimEnd(),
      type: 'double-quote'
    });
  }
}

console.log('\nPotential multiline string issues in generator template:', issues.length);
issues.slice(0, 20).forEach(x => {
  console.log('\nGenerator line ' + x.genLine + ' [' + x.type + ']:');
  console.log('  ' + x.content.substring(0, 100));
  console.log('  NEXT: ' + x.next.substring(0, 100));
});

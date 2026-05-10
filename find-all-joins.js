const fs = require('fs');
const gen = fs.readFileSync('scripts/generate-dashboard-html.js', 'utf8');

// Find ALL ].join( occurrences and determine if they're inside template literals
// Template 33 spans chars 23004 to 122383 (from previous analysis)
// Let me find all template boundaries more carefully

const templates = [];
let depth = 0;
let inStr = false;
let strChar = '';
let currentStart = -1;

for (let i = 0; i < gen.length; i++) {
  const c = gen[i];
  if ((inStr || depth > 0) && c === '\\') { i++; continue; }
  if (!inStr && depth === 0 && (c === "'" || c === '"')) { inStr = true; strChar = c; continue; }
  if (inStr && c === strChar) { inStr = false; continue; }
  if (inStr) continue;
  if (c === '`') {
    if (depth === 0) { currentStart = i; depth = 1; }
    else { templates.push({ start: currentStart, end: i }); depth = 0; currentStart = -1; }
  }
}

function isInTemplate(pos) {
  return templates.some(t => t.start < pos && pos < t.end);
}

// Find ALL ].join( occurrences
let pos = 0;
let joinCount = 0;
while ((pos = gen.indexOf('].join(', pos + 1)) !== -1) {
  joinCount++;
  console.log(`\nOccurrence ${joinCount} at char ${pos}:`);
  console.log('  In template:', isInTemplate(pos));
  console.log('  Context:', JSON.stringify(gen.substring(pos - 10, pos + 40)));
}

console.log('\nTotal ].join( occurrences:', joinCount);

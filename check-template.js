const fs = require('fs');
const gen = fs.readFileSync('scripts/generate-dashboard-html.js', 'utf8');

// Find all top-level template literal boundaries (not nested)
const starts = [];
const ends = [];
let depth = 0;
let inRegStr = false;
let strChar = '';

for (let i = 0; i < gen.length; i++) {
  const c = gen[i];
  
  // Handle escape sequences in strings/templates
  if ((inRegStr || depth > 0) && c === '\\') {
    i++; // skip next char
    continue;
  }
  
  // Handle regular string delimiters
  if (!inRegStr && depth === 0 && (c === "'" || c === '"')) {
    inRegStr = true;
    strChar = c;
    continue;
  }
  if (inRegStr && c === strChar) {
    inRegStr = false;
    continue;
  }
  if (inRegStr) continue;
  
  // Handle backticks
  if (c === '`') {
    if (depth === 0) {
      starts.push(i);
      depth = 1;
    } else {
      ends.push(i);
      depth = 0;
    }
  }
}

console.log('Template literals found:', starts.length);
starts.forEach((s, i) => {
  if (ends[i] !== undefined) {
    const len = ends[i] - s;
    console.log(`Template ${i + 1}: starts at ${s}, ends at ${ends[i]}, length ${len}`);
    // Check if boardNodeRow is inside
    const boardNodeRowPos = 77835;
    if (s < boardNodeRowPos && boardNodeRowPos < ends[i]) {
      console.log('  --> boardNodeRow IS inside this template');
    }
    // Check if rowTitle is inside
    const rowTitlePos = 78756;
    if (s < rowTitlePos && rowTitlePos < ends[i]) {
      console.log('  --> rowTitle IS inside this template');
    }
  }
});

// Also: find the actual text around char 78756 (rowTitle)
console.log('\nContext around rowTitle (char 78756):');
console.log(gen.substring(78700, 78800));

// Check what the generator file has around line 1560
// Count chars to get to line 1560
const genLines = gen.split('\n');
let charCount = 0;
for (let i = 0; i < 1559; i++) charCount += genLines[i].length + 1;
console.log('\nLine 1560 starts at char:', charCount);
console.log('Line 1560 content:', JSON.stringify(genLines[1559]));
console.log('Is it inside any template?');
for (let t = 0; t < starts.length; t++) {
  if (ends[t] !== undefined && starts[t] < charCount && charCount < ends[t]) {
    console.log('  YES - inside template', t + 1, '(chars', starts[t], '-', ends[t] + ')');
    break;
  }
}

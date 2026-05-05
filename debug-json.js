const fs = require('fs');
const html = fs.readFileSync('dashboard.html', 'utf8');

// Find the position 1107917 and show context
const pos = 1107917;
const start = pos - 100;
const end = pos + 100;
const context = html.substring(start, end);

console.log('Context around error position 1107917:');
console.log('---');
console.log(context);
console.log('---');
console.log('');

// Try to find DATA = 
const dataStart = html.indexOf('const DATA = ') + 13;
console.log('DATA starts at:', dataStart);

// Try to parse just the JSON portion
const sampleSize = 200000;
for (let endOffset = 1107900; endOffset <= 1107950; endOffset += 5) {
  const jsonCandidate = html.substring(dataStart, dataStart + endOffset);
  try {
    JSON.parse(jsonCandidate);
    console.log(`Valid JSON up to offset ${endOffset}`);
  } catch (e) {
    if (endOffset === 1107900) {
      console.log(`Invalid JSON at offset ${endOffset}: ${e.message.substring(0, 80)}`);
    }
  }
}

// Check what's actually at position 1107917
console.log('');
console.log('Character code at 1107917:', html.charCodeAt(1107917));
console.log('Substring 1107910-1107925:', JSON.stringify(html.substring(1107910, 1107925)));

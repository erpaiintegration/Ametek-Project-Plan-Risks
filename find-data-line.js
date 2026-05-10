const fs = require('fs');

const html = fs.readFileSync('./dashboard.html', 'utf8');
const lines = html.split('\n');

// Find the line with "const DATA ="
let dataLineNum = -1;
for (let i = 0; i < lines.length; i++) {
  if (lines[i].includes('const DATA =')) {
    dataLineNum = i;
    break;
  }
}

if (dataLineNum === -1) {
  console.log('ERROR: Could not find "const DATA =" line');
  process.exit(1);
}

console.log(`Found "const DATA =" at line ${dataLineNum + 1}`);
console.log(`\n--- Context: lines ${dataLineNum - 2} to ${dataLineNum + 25} ---\n`);

for (let i = Math.max(0, dataLineNum - 2); i < Math.min(lines.length, dataLineNum + 25); i++) {
  const lineContent = lines[i].length > 200 ? lines[i].substring(0, 200) + '...[truncated]' : lines[i];
  console.log(`${i + 1}: ${lineContent}`);
}

// Check if DATA line ends with semicolon
const dataLine = lines[dataLineNum];
console.log(`\n--- Data line analysis ---`);
console.log(`Length: ${dataLine.length}`);
console.log(`First 100 chars: ${dataLine.substring(0, 100)}`);
console.log(`Last 100 chars: ${dataLine.substring(dataLine.length - 100)}`);

// Find where the DATA object ends (look for matching closing brace/bracket)
const dataStart = dataLine.indexOf('{');
if (dataStart !== -1) {
  let braceCount = 0;
  let bracketCount = 0;
  let inDataObject = false;
  let dataEndPos = -1;
  
  for (let i = dataStart; i < dataLine.length; i++) {
    if (dataLine[i] === '{') braceCount++;
    else if (dataLine[i] === '}') {
      braceCount--;
      if (braceCount === 0 && bracketCount === 0) {
        dataEndPos = i;
        inDataObject = true;
        break;
      }
    }
    else if (dataLine[i] === '[') bracketCount++;
    else if (dataLine[i] === ']') bracketCount--;
  }
  
  if (dataEndPos !== -1) {
    console.log(`\nDATA object ends at position ${dataEndPos} (character: '${dataLine[dataEndPos]}')`);
    console.log(`Characters after DATA object end: '${dataLine.substring(dataEndPos, Math.min(dataEndPos + 10, dataLine.length))}'`);
  }
}

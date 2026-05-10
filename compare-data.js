const fs = require('fs');
const path = require('path');

// Load dashboard.html and extract the actual DATA
const html = fs.readFileSync('./dashboard.html', 'utf8');
const scriptStart = html.indexOf('<script>') + '<script>'.length;
const scriptEnd = html.indexOf('</script>');
const rawScript = html.substring(scriptStart, scriptEnd);

// Find where DATA ends
const dataEndPos = rawScript.indexOf('};') + 2;
const dataLineStr = rawScript.substring(12, dataEndPos); // Skip "const DATA = "
const actualData = JSON.parse(dataLineStr);

console.log('=== Actual DATA from dashboard.html ===\n');
console.log(`DATA keys: ${Object.keys(actualData).join(', ')}`);
console.log(`\n--- Size breakdown ---`);

for (const key of Object.keys(actualData)) {
  const strSize = JSON.stringify(actualData[key]).length;
  console.log(`${key}: ${strSize} bytes`);
  
  if (key === 'tasks') {
    console.log(`  - ${actualData[key].length} tasks`);
  } else if (Array.isArray(actualData[key])) {
    console.log(`  - ${actualData[key].length} items`);
  }
}

// Now check what the dashboard-data.json contains for full tasks
console.log('\n=== Full task data from dashboard-data.json ===\n');
const fullData = JSON.parse(fs.readFileSync('./dashboard-data.json', 'utf8'));
console.log(`Full data keys: ${Object.keys(fullData).join(', ')}`);
console.log(`Full tasks: ${fullData.tasks.length}`);

// Check if actionItems are in the actual DATA
if (actualData.actionItems) {
  console.log(`\n=== Action items in DATA ===`);
  console.log(`Count: ${actualData.actionItems.length}`);
  if (actualData.actionItems.length > 0) {
    console.log(`Sample item keys: ${Object.keys(actualData.actionItems[0]).join(', ')}`);
  }
}

// Check riskRecords
if (actualData.riskRecords) {
  console.log(`\n=== Risk records in DATA ===`);
  console.log(`Count: ${actualData.riskRecords.length}`);
  if (actualData.riskRecords.length > 0) {
    console.log(`Sample record keys: ${Object.keys(actualData.riskRecords[0]).join(', ')}`);
  }
}

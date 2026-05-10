const fs = require('fs');
const path = require('path');

// Load the actual generator to understand the structure
const DATA_FILE = path.join(__dirname, 'dashboard-data.json');
const data = fs.readFileSync(DATA_FILE, 'utf8');
const fullPayload = JSON.parse(data);

console.log(`Full task data loaded: ${fullPayload.tasks.length} tasks`);

// Simulate bootstrap selection (top 200)
const bootstrapTasks = fullPayload.tasks.slice(0, 200);
console.log(`Bootstrap tasks: ${bootstrapTasks.length}`);

// Create bootstrap payload like the generator does
const bootstrapPayload = {
  generatedAt: new Date().toISOString(),
  metrics: { /* mock */ },
  riskTypeBreakdown: [],
  statusBreakdown: [],
  topWorkstreams: [],
  riskRecords: [],
  actionItems: [],
  tasks: bootstrapTasks
};

// Stringify it like the generator does
const data_str = JSON.stringify(bootstrapPayload).replace(/<\/script>/gi, "<\\/script>");

console.log(`\nStringified data length: ${data_str.length}`);
console.log(`First 100 chars: ${data_str.substring(0, 100)}`);
console.log(`Last 100 chars: ${data_str.substring(data_str.length - 100)}`);

// Create the template line like the generator
const templateLine = `const DATA = ${data_str};`;
console.log(`\nTemplate line length: ${templateLine.length}`);
console.log(`Template line ends with: ${templateLine.substring(templateLine.length - 50)}`);

// Test if it parses
console.log(`\n--- Test parsing template line ---`);
try {
  new Function(templateLine);
  console.log('✓ Template line parses successfully!');
} catch (e) {
  console.log(`✗ Error: ${e.message}`);
}

// Test with the next variable declarations
const next2 = `const DATA = ${data_str};\nlet activeFilter = null;`;
console.log(`\n--- Test template + line 2 ---`);
try {
  new Function(next2);
  console.log('✓ Parses successfully!');
} catch (e) {
  console.log(`✗ Error: ${e.message}`);
}

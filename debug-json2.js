const fs = require('fs');
const html = fs.readFileSync('dashboard.html', 'utf8');

// Extract DATA object
const dataStart = html.indexOf('const DATA = ') + 13;
const script = html.substring(dataStart);

// Find where the JSON ends (look for }; followed by let)
const match = script.match(/^(\{[\s\S]*?\})\s*;/);
if (!match) {
  console.log('Could not find JSON object');
  process.exit(1);
}

const jsonStr = match[1];
console.log('Extracted JSON length:', jsonStr.length);

// Try parsing
try {
  const data = JSON.parse(jsonStr);
  console.log('✓ JSON is valid!');
  console.log('Tasks:', data.tasks.length);
  console.log('Risks:', data.riskRecords.length);
} catch (err) {
  console.log('✗ JSON parse error:', err.message);
  
  // Binary search to find where it breaks
  let left = 0, right = jsonStr.length;
  let lastValid = 0;
  
  while (left < right) {
    const mid = Math.floor((left + right) / 2);
    try {
      JSON.parse(jsonStr.substring(0, mid) + '}');  // Try closing early
      lastValid = mid;
      left = mid + 1;
    } catch {
      right = mid;
    }
  }
  
  console.log('Last valid JSON position:', lastValid);
  console.log('Context:', JSON.stringify(jsonStr.substring(lastValid - 100, lastValid + 100)));
}

const fs = require('fs');
const html = fs.readFileSync('dashboard.html', 'utf8');
const scriptStart = html.indexOf('<script>') + 8;
const scriptEnd = html.indexOf('</script>');
const script = html.slice(scriptStart, scriptEnd);

// Find where "const DATA = " starts
const dataStart = script.indexOf('const DATA = ');
const jsonStart = dataStart + 13; // after "const DATA = "

// Find the matching closing brace by counting braces
let braceCount = 0;
let inString = false;
let stringChar = '';
let jsonEnd = jsonStart;

for (let i = jsonStart; i < script.length; i++) {
  const char = script[i];
  const prevChar = script[i - 1];
  
  // Handle string state
  if ((char === '"' || char === "'" || char === '`') && prevChar !== '\\') {
    if (!inString) {
      inString = true;
      stringChar = char;
    } else if (char === stringChar) {
      inString = false;
    }
  }
  
  // Only count braces outside strings
  if (!inString) {
    if (char === '{') braceCount++;
    else if (char === '}') {
      braceCount--;
      if (braceCount === 0) {
        jsonEnd = i + 1;
        break;
      }
    }
  }
}

const jsonStr = script.slice(jsonStart, jsonEnd);
console.log('Extracted JSON length:', jsonStr.length);
console.log('Last 200 chars:', jsonStr.substring(jsonStr.length - 200));

console.log('\n--- Validating JSON ---');
try {
  const parsed = JSON.parse(jsonStr);
  console.log('✓ JSON is valid!');
  console.log('  Keys:', Object.keys(parsed));
  console.log('  Task count:', parsed.tasks?.length || 0);
} catch (e) {
  console.log('✗ JSON error:', e.message);
  const pos = parseInt(e.message.match(/position (\d+)/)?.[1] || 0);
  if (pos > 0) {
    console.log('  Error at position:', pos);
    console.log('  Context (pos-50 to pos+50):');
    console.log(jsonStr.substring(pos - 50, pos + 50));
  }
}

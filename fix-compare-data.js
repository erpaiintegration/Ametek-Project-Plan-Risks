const fs = require('fs');

// Load dashboard.html and extract the actual DATA
const html = fs.readFileSync('./dashboard.html', 'utf8');
const scriptStart = html.indexOf('<script>') + '<script>'.length;
const scriptEnd = html.indexOf('</script>');
const rawScript = html.substring(scriptStart, scriptEnd);

// Find where DATA ends - look for "};" pattern
const dataStartIdx = rawScript.indexOf('const DATA = ');
const objectStartIdx = rawScript.indexOf('{', dataStartIdx);
const semiIdx = rawScript.indexOf('};', objectStartIdx);

if (objectStartIdx < 0 || semiIdx < 0) {
  console.error('Could not find DATA object boundaries');
  process.exit(1);
}

// Extract just the JSON object
const dataLineStr = rawScript.substring(objectStartIdx, semiIdx + 1);
console.log(`DATA object size: ${dataLineStr.length} bytes`);
console.log(`First 100 chars: ${dataLineStr.substring(0, 100)}`);
console.log(`Last 100 chars: ${dataLineStr.substring(dataLineStr.length - 100)}`);

// Try to parse it
let actualData;
try {
  actualData = JSON.parse(dataLineStr);
  console.log(`\n✓ DATA object parsed successfully`);
} catch (e) {
  console.log(`✗ JSON parse error: ${e.message}`);
  // Show context around the error
  const pos = parseInt(e.message.match(/position (\d+)/)?.[1] || '0');
  if (pos > 0) {
    console.log(`\nError at position ${pos}:`);
    console.log(`...${dataLineStr.substring(Math.max(0, pos - 50), pos + 50)}...`);
  }
  process.exit(1);
}

// Print stats
console.log(`\nDATA keys: ${Object.keys(actualData).join(', ')}`);
console.log(`Tasks in DATA: ${actualData.tasks.length}`);

// Now test parsing the full script
console.log(`\n--- Testing full script parsing ---`);
try {
  new Function(rawScript);
  console.log('✓ Full script parses successfully!');
} catch (e) {
  console.log(`✗ ERROR: ${e.message}`);
}

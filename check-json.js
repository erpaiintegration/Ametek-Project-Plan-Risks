const fs = require('fs');
const html = fs.readFileSync('dashboard.html', 'utf-8');
const scriptStart = html.indexOf('<script') + html.substring(html.indexOf('<script')).indexOf('>') + 1;
const scriptEnd = html.indexOf('</script>', scriptStart);
const scriptContent = html.substring(scriptStart, scriptEnd);

// Extract just the JSON data
const dataStartIdx = scriptContent.indexOf('const DATA = ') + 'const DATA = '.length;
const dataEndIdx = scriptContent.indexOf('};', dataStartIdx) + 1;
const jsonString = scriptContent.substring(dataStartIdx, dataEndIdx);

console.log('Extracted JSON length:', jsonString.length);
console.log('First 150 chars:', jsonString.substring(0, 150));
console.log('Last 150 chars:', jsonString.substring(jsonString.length - 150));

try {
  JSON.parse(jsonString);
  console.log('\n✓ DATA JSON is valid!');
} catch (e) {
  console.log(`\n✗ JSON parse error: ${e.message}`);
  console.log('Error location:', e.stack.split('\n')[1]);
  
  // Try to pinpoint where the error is
  const lines = jsonString.split('\n');
  console.log(`\nTrying to parse line by line to find the error...`);
  let validUpTo = 0;
  for (let i = 0; i < lines.length; i++) {
    const partial = jsonString.substring(0, jsonString.indexOf(lines[i]) + lines[i].length);
    try {
      JSON.parse(partial);
      validUpTo = i;
    } catch (e2) {
      console.log(`Error first appears around line ${i}`);
      console.log(`Line ${i - 1}: ${lines[i-1]?.substring(0, 100) || 'N/A'}`);
      console.log(`Line ${i}: ${lines[i].substring(0, 100)}`);
      console.log(`Line ${i + 1}: ${lines[i+1]?.substring(0, 100) || 'N/A'}`);
      break;
    }
  }
}

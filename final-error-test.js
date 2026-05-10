const fs = require('fs');

const html = fs.readFileSync('./dashboard.html', 'utf8');

// Get the raw script content
const scriptStart = html.indexOf('<script>') + '<script>'.length;
const scriptEnd = html.indexOf('</script>');
const rawScript = html.substring(scriptStart, scriptEnd);

console.log(`=== Full script analysis ===`);
console.log(`Script length: ${rawScript.length}`);
console.log(`Script lines: ${rawScript.split('\n').length}`);

// Try to parse the full script
console.log(`\n--- Attempting to parse full script ---`);
try {
  new Function(rawScript);
  console.log('✓ FULL script parses successfully!');
} catch (e) {
  console.log(`✗ ERROR: ${e.message}`);
  
  // Binary search to find the exact error
  const lines = rawScript.split('\n');
  let low = 0;
  let high = lines.length;
  
  while (low < high) {
    const mid = Math.floor((low + high) / 2);
    const testCode = lines.slice(0, mid + 1).join('\n');
    
    try {
      new Function(testCode);
      low = mid + 1;
    } catch (err) {
      high = mid;
    }
  }
  
  console.log(`\n✗ First ${low} lines parse, error at line ${low + 1}`);
  console.log(`\nContext (lines ${Math.max(1, low - 2)} to ${low + 3}):`);
  for (let i = Math.max(0, low - 2); i < Math.min(lines.length, low + 3); i++) {
    const prefix = i === low ? '>>> ' : '    ';
    const line = lines[i].length > 120 ? lines[i].substring(0, 120) + '...' : lines[i];
    console.log(`${prefix}${i + 1}: ${line}`);
  }
  
  // Show the exact error
  try {
    new Function(lines.slice(0, low + 1).join('\n'));
  } catch (err2) {
    console.log(`\nExact error: ${err2.message}`);
  }
}

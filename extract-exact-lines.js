const fs = require('fs');
const html = fs.readFileSync('dashboard.html', 'utf8');
const scriptStart = html.indexOf('<script>') + 8;
const scriptEnd = html.indexOf('</script>');
const fullScript = html.slice(scriptStart, scriptEnd);
const lines = fullScript.split('\n');

// Get exact lines 1-22 (index 0-21)
const testCode = lines.slice(0, 22).join('\n');

console.log('Attempting to parse lines 1-22...');
console.log(`Code length: ${testCode.length}`);

try {
  new Function(testCode);
  console.log('✓ Valid!');
} catch (e) {
  console.log(`✗ Error: ${e.message}`);
  
  // Save to file for inspection
  fs.writeFileSync('problematic-code.txt', testCode, 'utf8');
  console.log('Saved problematic code to problematic-code.txt');
  
  // Try to narrow down by testing adding one more line at a time after line 21
  console.log('\n--- Testing incremental additions after line 21 ---');
  
  for (let i = 22; i < Math.min(30, lines.length); i++) {
    const incrementalCode = lines.slice(0, i + 1).join('\n');
    try {
      new Function(incrementalCode);
      console.log(`Lines 1-${i + 1}: ✓`);
    } catch (err) {
      console.log(`Lines 1-${i + 1}: ✗ ${err.message}`);
      break;
    }
  }
}

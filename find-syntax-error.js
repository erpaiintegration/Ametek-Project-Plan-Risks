const fs = require('fs');
const html = fs.readFileSync('docs/index.html', 'utf8');
const scriptStart = html.indexOf('<script>');
const scriptEnd = html.indexOf('</script>');
const script = html.substring(scriptStart + 8, scriptEnd);

// Try to find where error occurs
const lines = script.split('\n');
console.log(`Total lines: ${lines.length}`);

// Test if full script can run
try {
  new Function(script);
  console.log('✓ Full script parses');
} catch (e) {
  console.log(`✗ Full script error: ${e.message}`);
}

// Find the problematic line
let testCode = '';
for (let i = 0; i < Math.min(200, lines.length); i++) {
  testCode += lines[i] + '\n';
  try {
    new Function(testCode);
  } catch (e) {
    console.log(`\n✗ Error first appears when including line ${i}:`);
    console.log(`Line ${i}: ${lines[i].substring(0, 150)}`);
    console.log(`Error: ${e.message}`);
    
    // Show context
    console.log('\nContext (lines around error):');
    for (let j = Math.max(0, i-2); j <= Math.min(i+2, lines.length-1); j++) {
      console.log(`  Line ${j}: ${lines[j].substring(0, 100)}`);
    }
    break;
  }
}

const fs = require('fs');
const html = fs.readFileSync('dashboard.html', 'utf8');
const scriptStart = html.indexOf('<script>') + 8;
const scriptEnd = html.indexOf('</script>');
const script = html.slice(scriptStart, scriptEnd);
const lines = script.split('\n');

console.log('First 5 lines:');
for (let i = 0; i < 5; i++) {
  const preview = lines[i].substring(0, 60);
  console.log(`  Line ${i}: ${preview}...`);
}

console.log('\n=== TESTING CUMULATIVE (Lines 1-X starting from line 0) ===');
for (let i = 0; i <= 25; i++) {
  const code = lines.slice(0, i + 1).join('\n');
  try {
    new Function(code);
    console.log(`Lines 1-${i + 1}: OK`);
  } catch (e) {
    console.log(`\n!!! Lines 1-${i + 1}: ERROR !!!`);
    console.log(`  Error: ${e.message}`);
    console.log(`  Problem at line ${i + 1}: ${lines[i]}`);
    console.log('\nLast 3 lines:');
    for (let j = Math.max(0, i - 2); j <= i; j++) {
      console.log(`  Line ${j + 1}: ${lines[j]}`);
    }
    break;
  }
}

const fs = require('fs');
const html = fs.readFileSync('dashboard.html', 'utf8');
const scriptStart = html.indexOf('<script>') + 8;
const scriptEnd = html.indexOf('</script>');
const fullScript = html.slice(scriptStart, scriptEnd);
const lines = fullScript.split('\n');

console.log('Lines 1-22:');
for (let i = 0; i < 22; i++) {
  const line = lines[i];
  console.log(`${i + 1}: ${line.substring(0, 100)}`);
}

console.log('\n--- Testing lines 1-20 ---');
const code20 = lines.slice(0, 20).join('\n');
try {
  new Function(code20);
  console.log('✓ Lines 1-20 are valid');
} catch (e) {
  console.log('✗ Lines 1-20 error:', e.message);
}

console.log('\n--- Testing lines 1-21 ---');
const code21 = lines.slice(0, 21).join('\n');
try {
  new Function(code21);
  console.log('✓ Lines 1-21 are valid');
} catch (e) {
  console.log('✗ Lines 1-21 error:', e.message);
}

console.log('\n--- Testing lines 1-22 ---');
const code22 = lines.slice(0, 22).join('\n');
try {
  new Function(code22);
  console.log('✓ Lines 1-22 are valid');
} catch (e) {
  console.log('✗ Lines 1-22 error:', e.message);
}

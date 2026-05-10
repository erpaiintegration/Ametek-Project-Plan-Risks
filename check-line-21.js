const fs = require('fs');
const html = fs.readFileSync('dashboard.html', 'utf8');
const scriptStart = html.indexOf('<script>') + 8;
const scriptEnd = html.indexOf('</script>');
const script = html.slice(scriptStart, scriptEnd);
const lines = script.split('\n');

console.log('Line 20:', JSON.stringify(lines[19]));
console.log('Line 21:', JSON.stringify(lines[20]));
console.log('Line 22:', JSON.stringify(lines[21]));

console.log('\nLines 1-21 code:');
const code121 = lines.slice(0, 21).join('\n');
console.log(code121);
console.log('\n\nTesting lines 1-21:');
try {
  new Function(code121);
  console.log('VALID');
} catch (e) {
  console.log('ERROR:', e.message);
}

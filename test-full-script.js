const fs = require('fs');
const html = fs.readFileSync('dashboard.html', 'utf8');
const scriptStart = html.indexOf('<script>') + 8;
const scriptEnd = html.indexOf('</script>');
const script = html.slice(scriptStart, scriptEnd);

console.log('Testing full script...');
try {
  new Function(script);
  console.log('✓ VALID! Dashboard JavaScript is syntactically correct!');
} catch (e) {
  console.log('✗ ERROR:', e.message);
  console.log('Error at character position:', e.toString().match(/\d+/)?.[0]);
}

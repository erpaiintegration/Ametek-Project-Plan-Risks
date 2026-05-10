const fs = require('fs');
const html = fs.readFileSync('dashboard.html', 'utf8');
const scriptStart = html.indexOf('<script>') + 8;
const scriptEnd = html.indexOf('</script>');
const script = html.slice(scriptStart, scriptEnd);

// Extract the DATA part
const dataStart = script.indexOf('const DATA = ');
const endSemicolon = script.indexOf(';', dataStart);
const jsonStr = script.slice(dataStart + 13, endSemicolon);

console.log('Attempting to parse the JSON...');
try {
  const parsed = JSON.parse(jsonStr);
  console.log('✓ JSON is valid!');
  console.log('  Keys:', Object.keys(parsed));
  console.log('  Task count:', parsed.tasks?.length || 0);
} catch (e) {
  console.log('✗ JSON parse error:', e.message);
  const errorPos = parseInt(e.message.match(/position (\d+)/)?.[1] || 0);
  if (errorPos > 0) {
    console.log('  Error position:', errorPos);
    console.log('  Context:', jsonStr.substring(errorPos - 50, errorPos + 50));
    console.log('  Char at error:', jsonStr.charCodeAt(errorPos), jsonStr[errorPos]);
  }
}

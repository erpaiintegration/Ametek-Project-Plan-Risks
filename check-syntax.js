const fs = require('fs');
const html = fs.readFileSync('dashboard.html', 'utf-8');
const scriptStart = html.indexOf('<script') + html.substring(html.indexOf('<script')).indexOf('>') + 1;
const scriptEnd = html.indexOf('</script>', scriptStart);
const scriptContent = html.substring(scriptStart, scriptEnd);

// Show first 30 lines
const lines = scriptContent.split('\n').slice(0, 30);
console.log('First 30 lines of script:');
console.log('='.repeat(120));
lines.forEach((l, i) => {
  const truncated = l.length > 120 ? l.substring(0, 120) + '...' : l;
  console.log(`${String(i+1).padStart(3)}: ${truncated}`);
});
console.log('='.repeat(120));

// Try to validate
try {
  new Function(scriptContent);
  console.log('\n✓ JavaScript syntax is VALID');
} catch (e) {
  console.log(`\n✗ Syntax error: ${e.message}`);
  console.log(`\nTrying to parse DATA constant separately...`);
  const dataMatch = scriptContent.match(/^const DATA = (\{[\s\S]*?\n\};)/m);
  if (dataMatch) {
    console.log(`DATA constant found, length: ${dataMatch[1].length}`);
    try {
      JSON.parse(dataMatch[1].substring(0, dataMatch[1].length - 1)); // Remove trailing semicolon
      console.log('✓ DATA JSON is valid');
    } catch (e2) {
      console.log(`✗ DATA JSON error: ${e2.message}`);
    }
  } else {
    console.log('✗ Could not extract DATA constant');
  }
}

const fs = require('fs');
const html = fs.readFileSync('dashboard.html', 'utf8');
const scriptStart = html.indexOf('<script>') + 8;
const scriptEnd = html.indexOf('</script>');
const fullScript = html.slice(scriptStart, scriptEnd);

console.log('Full script length:', fullScript.length);
console.log('Testing full script syntax...\n');

try {
  new Function(fullScript);
  console.log('✓ JavaScript is syntactically valid!');
} catch (e) {
  console.log('✗ Syntax error:', e.message);
  
  // Try to find the error location
  const lines = fullScript.split('\n');
  console.log(`\nTotal lines: ${lines.length}`);
  
  // Binary search
  let low = 1, high = lines.length, errorLine = -1;
  
  while (low <= high) {
    const mid = Math.floor((low + high) / 2);
    const code = lines.slice(0, mid).join('\n');
    
    try {
      new Function(code);
      low = mid + 1;
    } catch (err) {
      errorLine = mid;
      high = mid - 1;
    }
  }
  
  if (errorLine > 0) {
    console.log(`\n✗ Error first appears at line ${errorLine}:`);
    console.log('Context:');
    for (let i = Math.max(0, errorLine - 5); i < Math.min(lines.length, errorLine + 3); i++) {
      const marker = i === errorLine - 1 ? '>>> ' : '    ';
      console.log(`${marker}${i + 1}: ${lines[i].substring(0, 100)}`);
    }
  }
}

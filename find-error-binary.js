const fs = require('fs');
const html = fs.readFileSync('dashboard.html', 'utf8');
const scriptStart = html.indexOf('<script>') + 8;
const scriptEnd = html.indexOf('</script>');
const script = html.slice(scriptStart, scriptEnd);
const lines = script.split('\n');

console.log(`Total lines: ${lines.length}`);

// Binary search for the error
let low = 1, high = lines.length;
let errorLine = -1;

while (low <= high) {
  const mid = Math.floor((low + high) / 2);
  const code = lines.slice(0, mid).join('\n');
  
  try {
    new Function(code);
    low = mid + 1;
  } catch (e) {
    errorLine = mid;
    high = mid - 1;
  }
}

if (errorLine > 0) {
  console.log(`\n✗ Syntax error first appears at line ${errorLine}:`);
  console.log(`  Content: ${lines[errorLine - 1]}`);
  
  // Show context
  console.log(`\n--- Context (lines ${Math.max(1, errorLine - 3)} to ${errorLine}) ---`);
  for (let i = Math.max(0, errorLine - 4); i < errorLine; i++) {
    const marker = i === errorLine - 1 ? '>>> ' : '    ';
    console.log(`${marker}${i + 1}: ${lines[i].substring(0, 80)}`);
  }
} else {
  console.log('✓ No error found - script is valid!');
}

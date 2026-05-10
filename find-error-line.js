const fs = require('fs');
const html = fs.readFileSync('dashboard.html', 'utf-8');
const scriptStart = html.indexOf('<script') + html.substring(html.indexOf('<script')).indexOf('>') + 1;
const scriptEnd = html.indexOf('</script>', scriptStart);
let scriptContent = html.substring(scriptStart, scriptEnd);

console.log('Script content length:', scriptContent.length);
console.log('First 100 chars:', scriptContent.substring(0, 100));

// Try to compile and get exact error location
try {
  new Function(scriptContent);
  console.log('✓ Script is valid');
} catch (e) {
  console.log(`✗ Syntax error: ${e.message}`);
  
  // Try to narrow down the error by parsing chunks
  console.log('\nAttempting to find error location...');
  
  // Split by newlines and try progressively
  const lines = scriptContent.split('\n');
  console.log(`Total lines: ${lines.length}`);
  
  // Binary search for the error
  let errorLine = -1;
  let low = 0;
  let high = lines.length;
  
  while (low < high) {
    const mid = Math.floor((low + high) / 2);
    const partial = lines.slice(0, mid).join('\n');
    try {
      new Function(partial);
      low = mid + 1;
    } catch (e2) {
      high = mid;
      errorLine = mid;
    }
  }
  
  console.log(`\nError is around line ${errorLine}`);
  console.log(`\nShowing lines ${Math.max(0, errorLine - 5)} to ${Math.min(lines.length, errorLine + 10)}:`);
  for (let i = Math.max(0, errorLine - 5); i < Math.min(lines.length, errorLine + 10); i++) {
    const prefix = i === errorLine ? '>>> ' : '    ';
    const line = lines[i].length > 100 ? lines[i].substring(0, 100) + '...' : lines[i];
    console.log(`${prefix}Line ${i + 1}: ${line}`);
  }
}

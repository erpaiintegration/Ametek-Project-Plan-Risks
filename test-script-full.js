const fs = require('fs');

const html = fs.readFileSync('./dashboard.html', 'utf8');
const lines = html.split('\n');

// Find the script tag
let scriptStart = -1;
for (let i = 0; i < lines.length; i++) {
  if (lines[i].includes('<script>')) {
    scriptStart = i;
    break;
  }
}

if (scriptStart === -1) {
  console.log('ERROR: Could not find <script> tag');
  process.exit(1);
}

// Find the closing script tag
let scriptEnd = -1;
for (let i = scriptStart; i < lines.length; i++) {
  if (lines[i].includes('</script>')) {
    scriptEnd = i;
    break;
  }
}

if (scriptEnd === -1) {
  console.log('ERROR: Could not find </script> closing tag');
  process.exit(1);
}

console.log(`Script found from line ${scriptStart + 1} to line ${scriptEnd + 1}`);

// Extract the script content (excluding <script> and </script> tags)
let scriptContent = lines.slice(scriptStart + 1, scriptEnd).join('\n');

// Try to parse it
console.log('\n--- Testing JavaScript syntax ---');
try {
  new Function(scriptContent);
  console.log('✓ JavaScript syntax is valid!');
} catch (err) {
  console.log(`✗ ERROR: ${err.message}`);
  console.log(`\nLine where error occurs: ${err.stack}`);
  
  // Try to find the problematic area
  const errorMatch = err.stack.match(/line (\d+)/i);
  if (errorMatch) {
    const errorLine = parseInt(errorMatch[1]);
    console.log(`\n--- Content around error line ${errorLine} ---`);
    const contentLines = scriptContent.split('\n');
    const start = Math.max(0, errorLine - 5);
    const end = Math.min(contentLines.length, errorLine + 5);
    for (let i = start; i < end; i++) {
      const marker = i === errorLine - 1 ? '>>> ' : '    ';
      const line = contentLines[i].length > 150 ? contentLines[i].substring(0, 150) + '...' : contentLines[i];
      console.log(`${marker}${i + 1}: ${line}`);
    }
  }
}

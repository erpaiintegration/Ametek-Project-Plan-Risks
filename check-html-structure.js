const fs = require('fs');

// Load dashboard.html
const html = fs.readFileSync('./dashboard.html', 'utf8');

// Find script tags
const scriptOpenTag = html.indexOf('<script>');
const scriptCloseTag = html.indexOf('</script>');

console.log(`Script open tag at: ${scriptOpenTag}`);
console.log(`Script close tag at: ${scriptCloseTag}`);

// Get the content between tags
const tagEndPos = scriptOpenTag + '<script>'.length;
const scriptContent = html.substring(tagEndPos, scriptCloseTag);

console.log(`Script content length: ${scriptContent.length}`);
console.log(`First 200 chars: ${scriptContent.substring(0, 200)}`);
console.log(`Char codes of first 20 chars: ${Array.from(scriptContent.substring(0, 20)).map(c => `${c}(${c.charCodeAt(0)})`).join(', ')}`);

// Check for newlines
const firstNewline = scriptContent.indexOf('\n');
console.log(`\nFirst newline at position: ${firstNewline}`);
console.log(`Content before first newline (${firstNewline} chars): ${JSON.stringify(scriptContent.substring(0, firstNewline))}`);

// Split by newlines and show first few lines
const lines = scriptContent.split('\n');
console.log(`\n=== First 10 lines ===`);
for (let i = 0; i < Math.min(10, lines.length); i++) {
  const line = lines[i];
  const display = line.length > 100 ? line.substring(0, 100) + '...' : line;
  console.log(`Line ${i + 1} (${line.length} chars): ${display}`);
}

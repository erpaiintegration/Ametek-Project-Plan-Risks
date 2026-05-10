const fs = require('fs');
const html = fs.readFileSync('dashboard.html', 'utf8');
const scriptStart = html.indexOf('<script>') + 8;
const scriptEnd = html.indexOf('</script>');
const script = html.slice(scriptStart, scriptEnd);

// Find the DATA object
const dataStart = script.indexOf('const DATA = {');
const dataEnd = script.indexOf('};') + 2;
const dataLine = script.slice(dataStart, dataEnd);

console.log('DATA object length:', dataLine.length);
console.log('First 500 chars:');
console.log(dataLine.substring(0, 500));
console.log('\n...\n');
console.log('Last 500 chars:');
console.log(dataLine.substring(dataLine.length - 500));

// Count tasks
const taskMatch = dataLine.match(/"tasks":\s*\[(.*?)\]\s*}/);
if (taskMatch) {
  const taskCount = (taskMatch[1].match(/,"id":/g) || []).length + 1;
  console.log('\n✓ Task count in DATA:', taskCount);
}

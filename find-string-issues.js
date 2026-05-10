// Find lines in dashboard.html script where a single/double quote string spans to next line
const fs = require('fs');
const html = fs.readFileSync('dashboard.html', 'utf8');
const scriptStart = html.indexOf('<script>');
const scriptEnd = html.lastIndexOf('</script>');
const script = html.substring(scriptStart + 8, scriptEnd);
const lines = script.split('\n');

const issues = [];
for (let i = 0; i < lines.length - 1; i++) {
  const line = lines[i];
  // Count unescaped single quotes (ignoring escaped ones)
  let singles = 0;
  for (let j = 0; j < line.length; j++) {
    if (line[j] === '\\') { j++; continue; }
    if (line[j] === "'") singles++;
  }
  if (singles % 2 !== 0) {
    issues.push({ line: i + 1, content: line.substring(0, 120), next: lines[i + 1].substring(0, 60) });
  }
}

console.log('Lines with unclosed single-quoted strings:', issues.length);
issues.forEach(x => {
  console.log('\nLine ' + x.line + ': ' + x.content);
  console.log('  NEXT: ' + x.next);
});

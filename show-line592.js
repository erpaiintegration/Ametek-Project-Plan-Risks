const fs = require('fs');
const html = fs.readFileSync('dashboard.html', 'utf8');
const scriptStart = html.indexOf('<script>');
const scriptEnd = html.lastIndexOf('</script>');
const script = html.substring(scriptStart + 8, scriptEnd);
const lines = script.split('\n');

// Show lines 590-595 with exact char codes
for (let i = 589; i <= 595 && i < lines.length; i++) {
  const line = lines[i];
  process.stdout.write('Line ' + (i+1) + ' (' + line.length + ' chars): ');
  for (let j = 0; j < line.length; j++) {
    const code = line.charCodeAt(j);
    if (code > 31 && code < 128) process.stdout.write(line[j]);
    else process.stdout.write('[' + code + ']');
  }
  console.log();
}

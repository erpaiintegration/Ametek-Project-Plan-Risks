const fs = require('fs');
const html = fs.readFileSync('dashboard.html', 'utf8');
const scriptStart = html.indexOf('<script>');
const scriptEnd = html.lastIndexOf('</script>');
const script = html.substring(scriptStart + 8, scriptEnd);

// Find ].join( in the script
const joinIdx = script.indexOf("].join(");
console.log('Found ].join( at position:', joinIdx);

if (joinIdx >= 0) {
  const context = script.substring(joinIdx, joinIdx + 30);
  console.log('Content after ].join(:');
  for (let i = 0; i < context.length; i++) {
    const code = context.charCodeAt(i);
    process.stdout.write(context[i] + '(' + code + ') ');
  }
  console.log();
}

// Also check what the generator actually produces for that line
// Find the line in generator source
const gen = fs.readFileSync('scripts/generate-dashboard-html.js', 'utf8');
const genLines = gen.split('\n');
for (let i = 0; i < genLines.length; i++) {
  if (genLines[i].includes('].join(')) {
    console.log('\nGenerator line', i + 1, ':', JSON.stringify(genLines[i]));
  }
}

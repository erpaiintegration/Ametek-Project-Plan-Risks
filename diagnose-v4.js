const fs = require('fs');
const vm = require('vm');
const html = fs.readFileSync('dashboard.html', 'utf8');
const scriptStart = html.indexOf('<script>');
const scriptEnd = html.lastIndexOf('</script>');
const script = html.substring(scriptStart + 8, scriptEnd);
// Replace the big DATA line with a simple stub so we don't get false positives
const lines = script.split('\n').map((l, i) => i === 1 ? 'const DATA = {};' : l);

console.log('Total lines:', lines.length);

// Find last line that parses (with leading lines only) 
// Use vm.Script - binary search for first failure
let lo = 1, hi = lines.length;

// First check: does all-lines-simplified parse?
try {
  new vm.Script(lines.join('\n'));
  console.log('Full simplified script: OK - no parse error!');
  process.exit(0);
} catch(e) {
  console.log('Full simplified script: FAIL -', e.message);
}

// Binary search: find the smallest prefix that fails
while (lo < hi) {
  const mid = Math.floor((lo + hi) / 2);
  try {
    new vm.Script(lines.slice(0, mid).join('\n'));
    lo = mid + 1;
  } catch (e) {
    // Is this a "real" error or "unexpected end of input"?
    if (e.message.includes('Unexpected end of input') || e.message.includes('Unexpected end of script')) {
      // This is just incomplete code, not what we want
      lo = mid + 1;
    } else {
      hi = mid;
    }
  }
}

console.log('Error first appears at line:', lo);
console.log('Context:');
for (let i = Math.max(0, lo - 4); i <= Math.min(lines.length - 1, lo + 3); i++) {
  const marker = i === lo - 1 ? '>>>' : '   ';
  console.log(marker, 'Line ' + (i + 1) + ':', lines[i].substring(0, 120));
}

// Show exact error for this section
try {
  new vm.Script(lines.slice(0, lo).join('\n'));
  console.log('Lines 1-' + lo + ': OK');
} catch(e) {
  console.log('Lines 1-' + lo + ' error:', e.message);
}

// Check for non-ASCII chars around problem area
const problemLine = lines[lo - 1];
console.log('\nProblem line char codes:');
for (let i = 0; i < Math.min(problemLine.length, 80); i++) {
  const code = problemLine.charCodeAt(i);
  if (code > 127) {
    process.stdout.write('[' + problemLine[i] + ':' + code + ']');
  } else {
    process.stdout.write(problemLine[i]);
  }
}
console.log();

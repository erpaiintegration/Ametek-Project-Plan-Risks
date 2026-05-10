const fs = require('fs');
const html = fs.readFileSync('dashboard.html', 'utf8');
const scriptStart = html.indexOf('<script>');
const scriptEnd = html.lastIndexOf('</script>');
const script = html.substring(scriptStart + 8, scriptEnd);
const lines = script.split('\n');
const dataLine = lines[1]; // line 2 is the DATA

// 1. Check for invalid JS escape sequences in the data
// JSON.stringify should escape everything, but let's verify
// Find all \X sequences where X is not a valid JSON/JS escape char
const invalidEscapes = [];
for (let i = 0; i < dataLine.length - 1; i++) {
  if (dataLine[i] === '\\') {
    const next = dataLine[i + 1];
    if (!'"\\/bfnrtu0123456789'.includes(next)) {
      invalidEscapes.push({ pos: i, seq: '\\' + next, context: dataLine.substring(i - 5, i + 10) });
    }
    i++; // skip next char
  }
}
console.log('Invalid escape sequences:', invalidEscapes.length);
if (invalidEscapes.length > 0) {
  console.log('First 5:', JSON.stringify(invalidEscapes.slice(0, 5)));
}

// 2. Check for \u sequences (should all be valid 4-hex-digit unicode)
const uEscapes = [];
for (let i = 0; i < dataLine.length - 5; i++) {
  if (dataLine[i] === '\\' && dataLine[i + 1] === 'u') {
    const hex = dataLine.substring(i + 2, i + 6);
    if (!/^[0-9a-fA-F]{4}$/.test(hex)) {
      uEscapes.push({ pos: i, seq: dataLine.substring(i, i + 6) });
    }
    i++;
  }
}
console.log('Invalid \\u sequences:', uEscapes.length);
if (uEscapes.length > 0) console.log('First 5:', uEscapes.slice(0, 5));

// 3. Check for </script in DATA
const scriptTagInData = dataLine.indexOf('<\/script>');
const scriptTag2 = dataLine.indexOf('<\\\/script>');
console.log('</script> in DATA:', scriptTagInData);
console.log('<\\/script> in DATA:', scriptTag2);

// 4. Check last chars of line 2
console.log('Last 10 chars of DATA line:', JSON.stringify(dataLine.slice(-10)));

// 5. Try parsing ONLY the DATA as a module (which supports async/await at top level)
const vm = require('vm');
try {
  new vm.Script('const DATA = 1;');
  console.log('Simple vm.Script: OK');
} catch(e) {
  console.log('Simple vm.Script: FAIL', e.message);
}

// 6. What's the actual error on the full script?
try {
  new vm.Script(script);
  console.log('Full script vm.Script: OK');
} catch(e) {
  console.log('Full script vm.Script error:', e.message);
  console.log('Error type:', e.constructor.name);
}

// 7. Does the script parse if we skip line 2 entirely (replace DATA with a small object)?
const scriptWithoutData = lines.slice(0).map((l, i) => i === 1 ? 'const DATA = {};' : l).join('\n');
try {
  new vm.Script(scriptWithoutData);
  console.log('Script with simplified DATA: OK');
} catch(e) {
  console.log('Script with simplified DATA error:', e.message);
}

console.log('Done.');

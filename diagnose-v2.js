const fs = require('fs');
const html = fs.readFileSync('dashboard.html', 'utf8');
const scriptStart = html.indexOf('<script>');
const scriptEnd = html.lastIndexOf('</script>');
const script = html.substring(scriptStart + 8, scriptEnd);
const lines = script.split('\n');
const dataLine = lines[1]; // line 2 is the DATA

console.log('Data line length:', dataLine.length);

// Check for /* or // comment openers NOT inside strings
let inStr = false;
let strChar = '';
let i = 0;
let foundComment = false;
while (i < dataLine.length) {
  const c = dataLine[i];
  if (!inStr && (c === '"' || c === "'")) {
    inStr = true;
    strChar = c;
    i++;
    continue;
  }
  if (inStr && c === strChar && dataLine[i - 1] !== '\\') {
    inStr = false;
    i++;
    continue;
  }
  if (!inStr && c === '/' && dataLine[i + 1] === '*') {
    console.log('FOUND /* at pos', i, 'context:', dataLine.substring(Math.max(0, i - 20), i + 40));
    foundComment = true;
    break;
  }
  if (!inStr && c === '/' && dataLine[i + 1] === '/') {
    console.log('FOUND // at pos', i, 'context:', dataLine.substring(Math.max(0, i - 20), i + 40));
    foundComment = true;
    break;
  }
  i++;
}
if (!foundComment) console.log('No unescaped comments found. inStr at end:', inStr);

// Check for invalid escape sequences in the DATA
const escapeMatches = dataLine.match(/\\[^"\\/bfnrtu0-9]/g);
if (escapeMatches) {
  console.log('Invalid escape sequences:', escapeMatches.slice(0, 10));
} else {
  console.log('No invalid escape sequences found');
}

// Check for any chars > 127 (Unicode) in DATA  
const highChars = [];
for (let j = 0; j < dataLine.length; j++) {
  if (dataLine.charCodeAt(j) > 127) {
    highChars.push({ pos: j, char: dataLine[j], code: dataLine.charCodeAt(j) });
    if (highChars.length >= 5) break;
  }
}
if (highChars.length > 0) {
  console.log('High unicode chars in DATA:', highChars);
} else {
  console.log('No chars > 127 in DATA');
}

// Test async function works in new Function
try {
  new Function('async function test() { return 1; }');
  console.log('async function in new Function: OK');
} catch (e) {
  console.log('async function in new Function: FAIL -', e.message);
}

// Test the actual problem: does lines 1-21 being valid + line 22 break it?
// What if we test lines 1-22 in vm.Script instead?
const vm = require('vm');
try {
  new vm.Script(lines.slice(0, 22).join('\n'));
  console.log('vm.Script lines 1-22: OK (incomplete but no error)');
} catch (e) {
  console.log('vm.Script lines 1-22 error:', e.message);
  if (e.lineNumber) console.log('  At line:', e.lineNumber);
}

// Try running a small test: does the DATA have an unclosed string?
// Test by adding a dummy close
const testCode = lines[1] + '\nconst x = 1;';
try {
  new Function(testCode);
  console.log('DATA + dummy line: OK');
} catch (e) {
  console.log('DATA + dummy line error:', e.message);
}

// What about the DATA ending with a backslash before newline?
const dataEnd = dataLine.slice(-5);
console.log('Last 5 chars of DATA line:', [...dataEnd].map(c => c + '(' + c.charCodeAt(0) + ')').join(' '));

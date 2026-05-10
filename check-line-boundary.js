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

const scriptLines = lines.slice(scriptStart + 1);
const line1 = scriptLines[0];
const line2 = scriptLines[1];

console.log('=== Checking line boundaries ===\n');

// Get raw content between script tags
const scriptTag = lines.slice(0, scriptStart + 1).join('\n') + '\n';
const scriptContent = html.substring(html.indexOf('<script>') + '<script>'.length);
const scriptEndPos = scriptContent.indexOf('</script>');
const rawScript = scriptContent.substring(0, scriptEndPos);

// Find where DATA ends in raw content
const dataEndMarker = '};';
const dataEndPos = rawScript.indexOf(dataEndMarker) + dataEndMarker.length;

console.log(`Position where DATA ends: ${dataEndPos}`);
console.log(`Characters at DATA end: ${JSON.stringify(rawScript.substring(dataEndPos - 5, dataEndPos + 10))}`);

// Check for newline after };
const afterData = rawScript.substring(dataEndPos, dataEndPos + 10);
console.log(`\nCharacters right after '};':`);
for (let i = 0; i < Math.min(10, afterData.length); i++) {
  const char = afterData[i];
  const code = char.charCodeAt(0);
  const display = char === '\n' ? '\\n' : char === '\r' ? '\\r' : char === '\t' ? '\\t' : char === ' ' ? '·' : char;
  console.log(`[${i}] '${display}' (code ${code})`);
}

// Test if manually constructing the code works
console.log(`\n=== Manual code construction ===`);
const reconstructed = line1 + '\n' + line2 + '\n' + scriptLines[2] + '\n' + scriptLines[3];
try {
  new Function(reconstructed);
  console.log('✓ Manually constructed lines 1-4 parse successfully');
} catch (e) {
  console.log(`✗ Error: ${e.message}`);
}

// Check what the full raw script looks like after DATA
const afterDataSample = rawScript.substring(dataEndPos, dataEndPos + 200);
console.log(`\n=== Raw script content after DATA end ===`);
console.log(afterDataSample);

// Check if there's a character we're not accounting for
console.log(`\n=== Character analysis after DATA end ===`);
const after = rawScript.substring(dataEndPos, Math.min(dataEndPos + 100, rawScript.length));
for (let i = 0; i < Math.min(50, after.length); i++) {
  const code = after.charCodeAt(i);
  if (code < 32 || code > 126) {
    const display = code === 10 ? '\\n' : code === 13 ? '\\r' : `\\x${code.toString(16)}`;
    console.log(`[${i}] NON-ASCII: '${display}' (code ${code})`);
  }
}

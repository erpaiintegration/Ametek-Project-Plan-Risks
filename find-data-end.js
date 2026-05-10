const fs = require('fs');
const html = fs.readFileSync('dashboard.html', 'utf-8');
const scriptStart = html.indexOf('<script') + html.substring(html.indexOf('<script')).indexOf('>') + 1;
const scriptEnd = html.indexOf('</script>', scriptStart);
const scriptContent = html.substring(scriptStart, scriptEnd);

// Find where DATA constant ends
const dataStartIdx = scriptContent.indexOf('const DATA = ');
if (dataStartIdx >= 0) {
  console.log('DATA starts at position:', dataStartIdx);
  // Check the first 500 chars after DATA =
  const startOfJSON = dataStartIdx + 'const DATA = '.length;
  console.log('\nFirst 200 chars of DATA value:');
  console.log(scriptContent.substring(startOfJSON, startOfJSON + 200));
  
  // Find where the JSON object closes
  let braceCount = 0;
  let inString = false;
  let escapeNext = false;
  let dataEndIdx = -1;
  
  for (let i = startOfJSON; i < Math.min(startOfJSON + 200000, scriptContent.length); i++) {
    const c = scriptContent[i];
    
    if (escapeNext) {
      escapeNext = false;
      continue;
    }
    
    if (c === '\\') {
      escapeNext = true;
      continue;
    }
    
    if (c === '"' && !inString) {
      inString = true;
      continue;
    }
    if (c === '"' && inString) {
      inString = false;
      continue;
    }
    
    if (!inString) {
      if (c === '{') braceCount++;
      if (c === '}') braceCount--;
      
      if (braceCount === 0 && c === ';') {
        dataEndIdx = i;
        break;
      }
    }
  }
  
  if (dataEndIdx > 0) {
    console.log('\nDATA ends at position:', dataEndIdx);
    const dataLen = dataEndIdx - startOfJSON;
    console.log('DATA JSON length:', dataLen);
    console.log('\nLast 200 chars of DATA:');
    console.log(scriptContent.substring(dataEndIdx - 200, dataEndIdx + 5));
    console.log('\nLine after DATA:', scriptContent.substring(dataEndIdx + 1, dataEndIdx + 50));
  } else {
    console.log('\nCould not find end of DATA within first 200k chars');
    // Try to find opening brace count
    let count = 0;
    for (let i = startOfJSON; i < Math.min(startOfJSON + 5000, scriptContent.length); i++) {
      if (scriptContent[i] === '{') count++;
      if (scriptContent[i] === '}') count--;
    }
    console.log('Brace count in first 5000 chars:', count);
  }
}

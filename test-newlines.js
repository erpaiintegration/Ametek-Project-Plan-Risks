// Test: check if template literals preserve newlines
const template = `<html>
<head>
<title>Test</title>
</head>
<body>
<p>Hello</p>
</body>
</html>`;

console.log('Template includes newlines:', template.includes('\n'));
console.log('Newline count:', (template.match(/\n/g) || []).length);
console.log('Template length:', template.length);
console.log('First 50 chars:', JSON.stringify(template.substring(0, 50)));
console.log('Template as is:', template.substring(0,100));

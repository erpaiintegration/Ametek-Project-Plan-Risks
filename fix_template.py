#!/usr/bin/env python3
import os

# Change to the correct directory
os.chdir('C:\\Users\\jsw73\\OneDrive\\Ametek Project Plan Risks')

with open('schedule_intelligence/build.js', 'r', encoding='utf-8', errors='replace') as f:
    content = f.read()

# Replace the incorrect </html>\; with </html>`; and closing }
# The backtick should be after </html> to close the template string
content = content.replace('</html>\\;', '</html>`;\n}')

with open('schedule_intelligence/build.js', 'w', encoding='utf-8') as f:
    f.write(content)

print('Fixed template closure successfully')
print('File saved')

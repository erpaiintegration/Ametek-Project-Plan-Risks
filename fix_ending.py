#!/usr/bin/env python3

import os
os.chdir('C:\\Users\\jsw73\\OneDrive\\Ametek Project Plan Risks')

with open('schedule_intelligence/build.js', 'r', encoding='utf-8', errors='replace') as f:
    content = f.read()

# The file should end with the main function closed and then main(); called
# Currently it ends with just }
# We need to make sure it ends with:\n}main();

# Remove any trailing closing brace(s) and whitespace
content = content.rstrip()
while content.endswith('}'):
    content = content[:-1].rstrip()

# Now append the proper ending
content += '\n}\n\nmain();\n'

with open('schedule_intelligence/build.js', 'w', encoding='utf-8') as f:
    f.write(content)

print('Fixed file ending')

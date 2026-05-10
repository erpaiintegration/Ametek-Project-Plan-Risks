#!/usr/bin/env python3

import os
os.chdir('C:\\Users\\jsw73\\OneDrive\\Ametek Project Plan Risks')

with open('schedule_intelligence/build.js', 'r', encoding='utf-8', errors='replace') as f:
    lines = f.readlines()

# Find and remove the duplicate/corrupted lines
cleaned_lines = []
skip_count = 0

for i, line in enumerate(lines):
    # Skip the corrupted duplicate lines after the console.log for CPM
    if skip_count > 0:
        skip_count -= 1
        continue
    
    if 'mermaid = buildMermaid(tasks)' in line:
        # Skip this line and the next 3 lines which contain duplicates
        skip_count = 3
        # Before skipping, make sure we have a closing brace
        if cleaned_lines[-1].strip() != '}':
            cleaned_lines.append('}\n')
        continue
    
    cleaned_lines.append(line)

# Write back
with open('schedule_intelligence/build.js', 'w', encoding='utf-8') as f:
    f.writelines(cleaned_lines)

print('Removed corrupted duplicate lines')

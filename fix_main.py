#!/usr/bin/env python3

import os
os.chdir('C:\\Users\\jsw73\\OneDrive\\Ametek Project Plan Risks')

with open('schedule_intelligence/build.js', 'r', encoding='utf-8', errors='replace') as f:
    content = f.read()

# Replace the broken section
old_section = """  const metrics = computeScheduleStats(tasks, cpmSummary);
}
    source: { file: path.basename(taskCsv) },"""

new_section = """  const metrics = computeScheduleStats(tasks, cpmSummary);
  const mermaid = buildMermaid(tasks);
  const laneAnchors = buildLaneAnchors(tasks);
  
  const payload = {
    generatedAt: new Date().toISOString(),
    source: { file: path.basename(taskCsv) },"""

content = content.replace(old_section, new_section)

with open('schedule_intelligence/build.js', 'w', encoding='utf-8') as f:
    f.write(content)

print('Fixed main function structure')

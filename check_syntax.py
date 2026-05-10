#!/usr/bin/env python3
import sys

with open('schedule_intelligence/build.js', 'r', encoding='utf-8', errors='replace') as f:
    lines = f.readlines()

print("=== SCRIPT TAGS ===")
for i, line in enumerate(lines):
    if '<script' in line or '</script>' in line:
        print(f'{i+1}: {line.rstrip()[:100]}')

print("\n=== LINES AROUND </html> ===")
for i, line in enumerate(lines):
    if '</html>' in line:
        for j in range(max(0, i-2), min(len(lines), i+5)):
            print(f'{j+1}: {lines[j].rstrip()[:100]}')
        break

print("\n=== END OF FILE ===")
for line in lines[-10:]:
    print(line.rstrip()[:100])

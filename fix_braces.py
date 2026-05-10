with open('schedule_intelligence/build.js', 'r', encoding='utf-8', errors='replace') as f:
    content = f.read()

# Find where main() is called
idx = content.rfind('main();')
if idx > 0:
    # Insert a closing brace before main();
    before = content[:idx].rstrip()
    after = content[idx:]
    
    # Count braces in the content before main()
    open_b = before.count('{')
    close_b = before.count('}')
    
    print(f'Before main(): {open_b} open, {close_b} close')
    print(f'Difference: {open_b - close_b}')
    
    # Add closing braces as needed
    missing = open_b - close_b
    if missing > 0:
        new_content = before + '\n' + ('}\n' * missing) + '\n' + after
        with open('schedule_intelligence/build.js', 'w', encoding='utf-8') as f:
            f.write(new_content)
        print(f'Added {missing} closing brace(s)')
    else:
        print('No missing braces')

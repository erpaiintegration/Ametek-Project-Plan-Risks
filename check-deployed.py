import urllib.request

url = 'https://jsw735.github.io/Ametek-Project-Plan-Risks/index.html'
try:
    response = urllib.request.urlopen(url)
    content = response.read().decode('utf-8')
    print(f"Content length: {len(content)} bytes")
    print(f"Status: {response.status}")
    
    # Check for DATA
    if 'const DATA = {' in content:
        print("✓ Found: const DATA = {")
        idx = content.find('const DATA = {')
        print(f"  Position: {idx}")
        print(f"  Sample: {content[idx:idx+100]}")
    else:
        print("✗ NOT FOUND: const DATA = {")
    
    # Check for init
    if 'async function init()' in content:
        print("✓ Found: async function init()")
    else:
        print("✗ NOT FOUND: async function init()")
    
    # Check for loadTaskData
    if 'async function loadTaskData()' in content:
        print("✓ Found: async function loadTaskData()")
    else:
        print("✗ NOT FOUND: async function loadTaskData()")
        
    # Check for script closing
    if '</script>' in content:
        print("✓ Found: </script> closing tag")
    else:
        print("✗ NOT FOUND: closing script tag")
        
    # Check what's near the script close
    script_close_idx = content.find('</script>')
    print(f"\nLast 200 chars before </script>:")
    print(content[script_close_idx-200:script_close_idx+20])
        
except Exception as e:
    print(f"Error: {e}")

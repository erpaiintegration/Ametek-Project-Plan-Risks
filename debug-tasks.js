const { fetchTasks, fetchRisks } = require('./src/notion/client.ts');

(async () => {
  try {
    const tasks = await fetchTasks();
    const riskRecords = await fetchRisks();
    
    // Try to stringify just the tasks
    const tasksJson = JSON.stringify(tasks);
    console.log('Tasks JSON valid, length:', tasksJson.length);
    
    // Check for problematic patterns
    const problematicChars = tasksJson.match(/[\x00-\x1f"'\\]/g);
    if (problematicChars) {
      console.log('Found potentially problematic chars:', new Set(problematicChars));
    }
    
    // Try to parse it back
    const parsed = JSON.parse(tasksJson);
    console.log('Parse successful, tasks count:', parsed.length);
    
    // Check a sample task
    if (tasks.length > 0) {
      const sample = tasks[0];
      console.log('Sample task keys:', Object.keys(sample));
      console.log('Sample task:', JSON.stringify(sample, null, 2).substring(0, 500));
    }
  } catch (err) {
    console.error('Error:', err.message);
  }
})();

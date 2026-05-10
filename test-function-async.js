console.log('=== Testing Function() constructor with async ===\n');

// Test 1: Simple async function in Function()
console.log('Test 1: Simple async function');
try {
  new Function('async function test() {}');
  console.log('✓ Works');
} catch (e) {
  console.log(`✗ Error: ${e.message}`);
}

// Test 2: Assignment then async function
console.log('\nTest 2: Assignment then async function');
try {
  new Function('const DATA = {}; async function test() {}');
  console.log('✓ Works');
} catch (e) {
  console.log(`✗ Error: ${e.message}`);
}

// Test 3: Long JSON assignment then async
console.log('\nTest 3: Long JSON then async function');
try {
  const longJson = JSON.stringify({tasks: Array(1000).fill({id: '123', name: 'test'})});
  new Function(`const DATA = ${longJson}; async function test() {}`);
  console.log('✓ Works');
} catch (e) {
  console.log(`✗ Error: ${e.message}`);
}

// Test 4: Try-catch inside async in Function()
console.log('\nTest 4: Async function with try-catch in Function()');
try {
  new Function(`
    async function test() {
      try {
        await fetch('test');
      } catch (e) {
        console.log(e);
      }
    }
  `);
  console.log('✓ Works');
} catch (e) {
  console.log(`✗ Error: ${e.message}`);
}

// Test 5: Multiple statements
console.log('\nTest 5: Multiple statements with async');
try {
  new Function(`
    const DATA = {};
    let x = null;
    async function loadData() {
      try {
        const response = await fetch('test');
      } catch (e) {
      }
    }
  `);
  console.log('✓ Works');
} catch (e) {
  console.log(`✗ Error: ${e.message}`);
}

// Test 6: Exactly like our case
console.log('\nTest 6: Exactly matching our case structure');
try {
  const code = `const DATA = {"test": "value"};
let activeFilter = null;
let selectedTaskId = null;
let ganttRendered = false;
let planRendered = false;
let boardRendered = false;
let taskById = new Map();
let taskByUid = new Map();
let planState = null;
let planGanttTimer = null;
const boardState = { selectedActionId: null, selectedNodeId: null, expanded: new Set(), tree: null, nodeMap: new Map(), rootTask: null, criticalOnly: false, drivingOnly: false };
const fmt = d => d ? new Date(d).toLocaleDateString("en-US",{month:"short",day:"numeric",year:"2-digit"}) : "—";
const fmtShort = d => d ? new Date(d).toLocaleDateString("en-US",{month:"short",day:"numeric"}) : "—";
const esc = s => String(s||"").replace(/&/g,"&amp;").replace(/</g,"&lt;").replace(/>/g,"&gt;");
const trunc = (s,n) => s && s.length>n ? s.slice(0,n)+"…" : (s||"");
const riskKey = r => (r.type || "Risk") + " — " + (r.category || "(Uncategorized)");
const topRiskKey = task => task.linkedRisks.length ? riskKey(task.linkedRisks[0]) : "No linked risk";
const PERF = { initialTaskRows: 180, maxTaskRows: 320 };
let atRiskTasksCache = null;

async function loadTaskData() {
  try {
    const response = await fetch('test');
  } catch (e) {
  }
}`;
  
  new Function(code);
  console.log('✓ Works');
} catch (e) {
  console.log(`✗ Error: ${e.message}`);
}

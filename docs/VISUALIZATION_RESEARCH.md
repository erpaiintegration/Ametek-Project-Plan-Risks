# Milestone Report Visualization Research

**Date**: May 10, 2026  
**Purpose**: Evaluate modern Excel visualization technologies for PMO Huddle Report (milestone tracking with cascading dependencies)  
**Decision**: Three-tier approach combining Excel native features + Python visualization libraries

---

## Project Requirements

- Display milestones organized by ITC cycle (ITC1, ITC2, ITC3, UAT, Go-Live)
- Show hierarchical WBS structure with parent-child relationships
- Visualize cascading dependencies (when ITC2 is late, mark ITC3 as at-risk)
- Display KPIs: % Complete, Open Tasks, Blocking Status, Critical Path
- Enable filtering by ITC, status, or at-risk flag
- Support dynamic refresh from project plan data

---

## Technology Options Evaluated

### 1. Excel "Containers" Feature ⭐ Limited
**Status**: Available in Excel 365 and newer desktop versions (2021+)  
**What it does**: Groups shapes, charts, and text boxes as a single unit for coordinated formatting  
**Pros**:
- Clean visual grouping of milestone sections
- Can lock/move related objects together

**Cons**:
- Not available in older Excel versions
- Limited interactivity (no dependency visualization)
- Basic grouping only

**Verdict**: Good for UI organization, insufficient for dependency tracking

---

### 2. Modern Excel Chart Types ✅ Recommended
**Available**: Excel 2016+  

#### Sunburst Chart ⭐ Best for WBS Hierarchy
Shows hierarchical milestone structure with % complete as ring size
```
Center: Project
Ring 1: ITC1, ITC2, ITC3, UAT, Go-Live
Ring 2: Milestones per ITC
Ring 3: Child tasks (optional drill-down)
```

**Pros**:
- Native Excel feature, no add-ins needed
- Click to zoom into hierarchy levels
- Color-code by status (complete/in-progress/not-started)
- % complete controls ring thickness

**Cons**:
- Cannot show explicit dependencies (ITC2→ITC3 links)
- Limited to radial layout

#### Treemap Chart
Shows milestone breakdown by workstream with % complete as box size

**Pros**:
- Space-efficient hierarchy display
- Color by status, size by % complete
- Interactive drill-down

**Cons**:
- No dependency arrows
- Harder to see timeline/sequence

#### Timeline Slicer
Filter and view milestones by date range dynamically

**Pros**:
- Interactive date-based filtering
- Linked to all charts/tables on sheet

**Cons**:
- Cannot show cascading delays
- Requires data to be in Excel range (not live from PQ)

---

### 3. Python Visualization Libraries 🐍 Recommended

#### **Pyvis (networkx + interactive HTML)** ⭐ BEST FOR DEPENDENCIES
Generates interactive network diagram showing milestone nodes and dependency edges

**How it works**:
1. Build directed graph: nodes = milestones, edges = dependencies
2. Assign attributes: node color by status, size by % complete, label with name
3. Generate interactive HTML visualization
4. Export as PNG for embedding in Excel OR link as companion dashboard

**Example output**:
```
ITC1 Kickoff ────→ ITC1 Planning ────→ ITC1 Configuration
                                             ↓
                                    ITC1 Data Migration
                                             ↓
                                      ITC2 Planning  (BLOCKED if ITC1 late)
```

**Pros**:
- ⭐ Perfect for cascading dependencies
- ⭐ Shows downstream impact visually
- Interactive: click nodes to highlight paths, hover for details
- Can be embedded as PNG in Excel OR viewed as separate HTML dashboard
- Automatically detects critical path

**Cons**:
- Requires Python environment
- HTML output (separate from Excel) unless embedded as image

**Code sketch**:
```python
from pyvis.network import Network
import networkx as nx

# Build graph
G = nx.DiGraph()
for milestone in milestones:
    G.add_node(milestone['name'], 
               itc=milestone['itc'],
               pct=milestone['pct_complete'],
               status=milestone['status'])

for dep in dependencies:
    G.add_edge(dep['from'], dep['to'], relationship='prerequisite')

# Visualize
net = Network(directed=True, physics=True, height='750px', width='100%')
net.from_nx(G)
net.show('milestone_dependencies.html')
```

#### **Plotly (interactive Gantt + charts)**
Generates interactive timeline chart showing milestones, start/finish dates, % complete bars, and dependency connectors

**Pros**:
- Beautiful, professional appearance
- Hover shows all details (dates, owner, open tasks, blocking status)
- Can export to static PNG for Excel embedding
- HTML output for interactive companion dashboard

**Cons**:
- Dependency lines are less visually prominent than Pyvis
- Requires more data prep

**Output types**:
- Static PNG: Embed in Excel sheet
- Interactive HTML: Link in Excel or separate dashboard

#### **openpyxl + Matplotlib + Screenshots**
Generate custom Python visualizations and embed as images in Excel

**Pros**:
- Full control over visualization design
- Can create completely custom layouts
- Embeds directly into .xlsx

**Cons**:
- Static images (no interactivity)
- More complex code

---

### 4. Conditional Formatting + Data Validation ✅ Easy Built-in
Excel native feature for color/icon coding

**Examples**:
```excel
Icon Sets: ● (Complete), ◐ (In Progress), ○ (Not Started)
Data Bars: Show % complete as bar fill
Color Scales: Green → Amber → Red for at-risk status
Custom Rules: =IF(AND(BlockedByLate, Status="Future"), TRUE, FALSE)
```

**Pros**:
- Immediate visual feedback
- No add-ins or Python needed
- Works in all Excel versions

**Cons**:
- Cannot show connections between milestones
- Limited to single-cell context

---

## Recommended Implementation

### **Three-Tier Approach** (Best of All Worlds)

#### **Tier 1: Primary Excel Sheet** (In metrics.xlsx)
What users see when they open the file

**Components**:
- **Spill formula** (already have): Pulls milestone data from Power Query with live refresh
- **Sunburst or Treemap chart**: Displays WBS hierarchy with % complete
- **Conditional formatting**: Status icons, blocking indicators, at-risk highlighting
- **Slicer**: Filter by ITC or status
- **Data table**: Detailed milestone list with hyperlinks to dependencies

**Layout** (following PMO Huddle guidelines):
```
┌─────────────────────────────────────────────────┐
│  Milestone Summary Dashboard                    │
│  Overall: X% Complete | Y Open Tasks | Z Late  │
├─────────────────────────────────────────────────┤
│  [Sunburst Chart]          │  [Milestone Table] │
│   (WBS hierarchy)          │   Name|%|Tasks|... │
│                            │   ────────────────│
├─────────────────────────────────────────────────┤
│ ITC1: [Chart]  ITC2: [Chart]  ITC3: [Chart]    │
│ UAT: [Chart]   Go-Live: [Chart]                │
├─────────────────────────────────────────────────┤
│ [ITC Slicer]  [Status Filter]  [Date Slicer]  │
└─────────────────────────────────────────────────┘
```

**Refresh behavior**: Updates on every daily report refresh via Power Query

---

#### **Tier 2: Companion HTML Dashboard** (Optional, separate file)
For PMO team who want interactive exploration

**Contents**:
- **Pyvis dependency network**: Click to explore cascade effects
- **Plotly Gantt chart**: Interactive timeline with hover details
- **Metrics widgets**: Match Excel KPIs
- **Filtering**: By ITC, status, critical path

**File**: `docs/milestone-dependencies.html` (linked from Excel)  
**Update**: Generated nightly by Python script alongside daily report

---

#### **Tier 3: Embedded PNG Images** (Optional)
For presentation/reports that need static snapshot

**Contents**:
- Pyvis network diagram (exported as PNG)
- Plotly Gantt (exported as PNG)
- Embedded in Excel sheet or printed to PDF

---

## Implementation Sequence

### **Phase 1: Excel Native** (Week 1)
1. Create `PQ_MilestoneShow` query: Filter milestones (Flag9="Yes"), project WBS structure
2. Create `PQ_MilestoneTree` query: Add ITC classification, hierarchy order, status
3. Create `Milestone_Summary` sheet with:
   - Spill formula pulling `PQ_MilestoneTree`
   - Sunburst chart showing WBS + % complete
   - Conditional formatting for blocking status
   - Slicer for ITC filtering

### **Phase 2: Python Visualization** (Week 2)
1. Create `generate_milestone_dependencies.py`:
   - Read milestones and dependencies from daily report workbook
   - Build Pyvis network diagram
   - Generate interactive HTML
   - Export PNG for embedding
2. Integrate into daily report pipeline: `scripts/build_daily_report.py`
3. Generate companion HTML dashboard

### **Phase 3: Polish & Validation** (Week 3)
1. Test cascading delay detection accuracy
2. Validate visual clarity and usability
3. Create PDF export template (static snapshots)
4. Document for PMO team

---

## Technology Stack Summary

| Component | Technology | Language | Effort |
|-----------|-----------|----------|--------|
| **Milestone display** | Excel + Power Query | M/DAX | ⭐ Easy |
| **WBS visualization** | Excel Sunburst chart | Native | ⭐ Easy |
| **Status indicators** | Conditional formatting | Native | ⭐ Easy |
| **Dependency network** | Pyvis + networkx | Python | ⭐⭐ Medium |
| **Interactive timeline** | Plotly | Python | ⭐⭐ Medium |
| **Integration** | Daily report pipeline | Python | ⭐⭐ Medium |

---

## Key Benefits of This Approach

✅ **Familiar**: Excel sheet interface (no learning curve)  
✅ **Live**: Power Query refresh on daily schedule  
✅ **Visual**: Sunburst hierarchy + icons + colors  
✅ **Interactive** (opt-in): HTML companion for exploration  
✅ **Scalable**: Works with 50+ milestones without performance issues  
✅ **Professional**: Meets PMO Huddle Report standards  
✅ **Extensible**: Can add more visualizations later (network costs, resource allocation, etc.)

---

## Risks & Mitigations

| Risk | Mitigation |
|------|-----------|
| Sunburst chart gets cluttered with 50+ milestones | Use ITC slicer to filter; zoom into subtrees |
| Dependency graph too complex | Use layout algorithm (Spring, Hierarchical) to reduce edge crossing |
| Python updates break Pyvis HTML output | Pin package versions in requirements.txt |
| Cascading delay detection misses edge cases | Add unit tests for delay propagation logic |

---

## Next Steps

1. **Get approval** on three-tier approach
2. **Build Phase 1**: Create Power Query milestone queries + Sunburst chart
3. **Test cascading logic**: Verify delay detection works correctly
4. **Build Phase 2**: Pyvis network + Plotly Gantt
5. **Integrate**: Add to daily report pipeline
6. **Document**: Create user guide for PMO team

---

## Decision Record

**Chosen approach**: Three-tier (Excel + Python + HTML)  
**Date**: May 10, 2026  
**Rationale**: Balances simplicity (Excel sheet), technical capability (Python libs), and visual impact (dependency networks)  
**Alternative considered**: Full custom HTML dashboard — rejected as too much context switching from Excel  
**Alternative considered**: Only Pyvis network — rejected as lacks hierarchical drill-down of Sunburst  


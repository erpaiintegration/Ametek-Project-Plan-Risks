# PMO Dashboard Design Contract (Excel/OpenPyXL)

This contract defines the **required visual framework, layout system, and implementation steps** for all domain dashboards (Testing/Defects, Project Plan, RAID, Actions, etc.).

## 1) Purpose

Ensure every dashboard tab is:

- visually consistent (brand, typography, panel system)
- readable in leadership meetings and screenshots
- predictable to generate and maintain in code
- reusable across domains with minimal custom code

---

## 2) Governance Boundary (must follow)

- **Metrics workbook owns KPI logic** and canonical field definitions.
- **Dashboard/huddle workbook owns presentation only**:
  - grouping for display
  - top-N selection
  - chart helper tables
  - visual formatting/layout
- Domain dashboards **must not redefine KPI semantics** in display scripts.

---

## 3) Required Style Tokens

Each domain dashboard script must define and use tokenized constants (not ad-hoc literals):

- Color tokens (`COLOR_*`): canvas, panel, panel edge, panel shadows, text, semantic status colors.
- Typography tokens:
  - `FONT_FAMILY`
  - title/subtitle/section/table/card sizes
- Border tokens:
  - panel edge
  - section separator
  - table grid

### Non-negotiable

- Do not hardcode random border colors in cell styling.
- Do not mix many font families.
- Keep all style values centralized at top of script.

---

## 4) Required Layout System

Every domain sheet must use 3 primary zones (panelized):

1. **Left panel**: KPI cards (stacked)
2. **Center panel**: detail tables (stacked)
3. **Right panel**: charts and chart subtitles

### Required visual mechanics

- Canvas background tint behind the whole dashboard.
- Raised panel effect (surface + shadow + border).
- Rounded-corner illusion via corner cutback to canvas.
- Section labels at top of each panel.

---

## 5) Required Components

## 5.1 KPI cards (left rail)

Must include:

- label row + value row
- consistent spacing rhythm (card block height)
- semantic color coding only for status-relevant cards
- left alignment for quick scan

## 5.2 Center tables

Must include:

- section header bar
- column header styling
- zebra striping for body rows
- wrapped text in long columns
- top-N truncation where needed for readability

## 5.3 Right-side charts

Must include:

- chart title hierarchy prefix (`A.`, `B.`, `C.`)
- subtitle line describing scope (under/near chart title area)
- reduced clutter (drop non-essential legend/series)
- shortened long category labels for axis readability

---

## 6) Required Build Pattern (exact implementation flow)

Each new domain dashboard script must implement this order:

1. **Read canonical metrics tables** from metrics workbook.
2. **Compute display-only subsets**:
   - today/period slices
   - top-N lists
   - exceptions table
   - chart helper datasets
3. **Initialize workbook + sheet**.
4. **Apply dashboard canvas + column widths + row rhythm**.
5. **Draw raised panels** (left/center/right).
6. **Add panel labels**.
7. **Render title/subtitle**.
8. **Render left KPI cards**.
9. **Render center tables**.
10. **Write hidden/narrow helper table area for charts**.
11. **Add charts + chart subtitles**.
12. **Add transparent raw data sheet for drill-through**.
13. **Save workbook**.

---

## 7) Function Contract (recommended reusable API)

All domain dashboard scripts should expose these functions (or equivalent):

- `read_metrics_tables(...)`
- `format_dashboard_sheet(...)`
- `draw_raised_panel(...)`
- `add_panel_label(...)`
- `write_cards(...)`
- `write_table(...)`
- `add_charts(...)`
- `add_chart_subtitles(...)`
- `build_<domain>_report(...)`

This creates consistency and makes cross-domain review easier.

---

## 8) Domain Onboarding Template

For each new domain (example: RAID):

1. Add canonical PQ tables in metrics builder:
   - `PQ_RAID_Detail`, `PQ_RAID_Summary`, `PQ_RAID_Workstream` (or domain equivalent)
2. Add domain dashboard script:
   - `scripts/build_pmo_huddle_<domain>.py`
3. Use this contract’s style/layout primitives.
4. Output sheet name should be domain-specific but style-consistent.
5. Add command alias in `package.json`.
6. Add README section with run command and expected output path.

---

## 9) Gotchas and Quirks (must consider)

## 9.1 Excel file locks (Windows)

- If workbook is open in Excel, save may fail with `PermissionError`.
- Always close workbook before running generator (or write to alternate filename for testing).

## 9.2 Merged cells

- Writing into non-top-left merged cells causes openpyxl errors (`MergedCell` read-only).
- Keep helper/chart data far away from merged title/header zones.

## 9.3 Long labels

- Workstream/task text can break chart readability.
- Always shorten category labels for chart axes (`shorten_label`).

## 9.4 Over-plotting

- Too many categories or series reduces clarity.
- Use top-N and remove low-value series (example: keep Tests/Passed/Failed trend, omit blocked if noisy).

## 9.5 Date parsing variability

- Source date columns may infer inconsistently.
- Prefer explicit parse strategy where feasible; handle `NaT` safely.

## 9.6 Visual drift

- Avoid introducing ad-hoc color/font values in one domain.
- Any design token change should be reflected across all domain scripts.

## 9.7 Performance

- Large row-level sheets can slow formatting.
- Limit display tables (`head(N)`), keep raw data on separate sheet.

---

## 10) Review Checklist (definition of done)

A domain dashboard is compliant only if:

- [ ] Uses tokenized colors/fonts/borders.
- [ ] Contains left KPI, center tables, right charts panel structure.
- [ ] Includes panel labels and chart subtitles.
- [ ] Uses top-N and shortened chart labels.
- [ ] Includes a drill-through data sheet.
- [ ] Builds successfully with no errors.
- [ ] Visuals are readable at 100% zoom in one screen capture.

---

## 11) Current Reference Implementation

Primary reference:

- `scripts/build_pmo_huddle_report.py`

This script is the baseline for all future domains unless a documented exception is approved.

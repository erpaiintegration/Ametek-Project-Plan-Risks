# Ametek Project Plan Risks

Notion automation for importing MS Project CSV task data, mapping it into the Tasks database, and maintaining hierarchy/dependency relationships for management reporting.

## Import workflow

1. Put the latest MSP CSV file in `imports/staging/`.
2. Run `npm run import`.
3. Importer will:
   - upsert tasks by `Unique ID`
  - merge `Resource Names`, `Business Validation Owner`, and `Workstream` from `Project plan with resources and busines.csv` when present
  - sync `Team Member` table from `Resource table.csv` and task name usage
  - classify team members as `BPO` (if in Business Validation) or `Func/Tech` (if in Resources)
   - overwrite MSP-owned fields
   - keep Notion-owned execution fields untouched
  - set `Parent Task`, `Predecessor Tasks`, `Successor Tasks`, `Projects`, `Assignee Team Members`, and `Business Validation Team Members`
   - stamp audit fields per row:
     - `MSP Source File`
     - `MSP Import Version`
     - `MSP Imported At`
   - move the processed file to `imports/archive/`

## Commands

- `npm run import` (imports newest CSV in `imports/staging/`)
- `npm run import -- --file "path/to/file.csv"` (imports a specific file)
- `npm run report:daily` (builds the `Ametek SAP S4 Impl Daily Report` workbook from the configured nightly source files)
- `npm run link:context` (links all Tasks to project context pages/databases such as Dashboard, Project Status, and Risks & Issues)
- `npm run sync:risks` (creates or updates aggregate Risks & Issues records from slipping tasks and links them to Tasks)
- `npm run dashboard:html` (generates `dashboard.html` in the project root)
- `npm run dashboard:pages` (generates `dashboard.html` and copies it to `docs/index.html` for GitHub Pages)
- `npm run build`

## Ametek SAP S4 Impl Daily Report

The daily report pipeline consolidates multiple daily-updated files into a single Excel workbook named `Ametek_SAP_S4_Impl_Daily_Report.xlsx`.

What it does:

- reads separate project plan, testing, defects, and RAID files (CSV or Excel)
- profiles each file structure and maps columns into a common model
- builds relationship tabs using shared IDs/task names where available
- creates a workstream rollup to relate sources even when structures differ
- writes a metrics summary, change summary, source profile, relationship tabs, and raw normalized tabs into one central workbook
- stores a lightweight snapshot JSON so the next nightly run can summarize adds/changes/removals

Configuration:

- source mappings live in `config/daily-report-sources.example.json`
- default output folder is `outputs/daily_report/`
- environment placeholders are in `.env`:
  - `DAILY_REPORT_CONFIG`
  - `DAILY_REPORT_OUTPUT_DIR`

Nightly run options:

- manual: `npm run report:daily`
- scheduled: use Windows Task Scheduler to call `scripts/run-daily-report.ps1` after your daily source files land in the staging folder

Outputs:

- `outputs/daily_report/Ametek_SAP_S4_Impl_Daily_Report.xlsx`
- `outputs/daily_report/Ametek_SAP_S4_Impl_Daily_Report_summary.md`
- `outputs/daily_report/last_snapshot.json`

## Publish dashboard with GitHub Pages

1. Run `npm run dashboard:pages`.
2. Publish this folder to GitHub.
3. In GitHub repo settings, enable **Pages** and set the source to:
  - **Branch:** `main`
  - **Folder:** `/docs`
4. GitHub will publish the dashboard from `docs/index.html`.
5. Paste the resulting Pages URL into a Notion `/embed` block.

Every time the underlying Notion data changes, run `npm run dashboard:pages` again and push the updated `docs/index.html`.

## Required `.env` settings

- `NOTION_API_KEY`
- `NOTION_TASKS_DB_ID`

Optional but recommended:

- `NOTION_PROJECTS_DB_ID`
- `NOTION_AMETEK_PROJECT_PAGE_ID`
- `NOTION_TEAM_MEMBERS_DB_ID` (auto-discovered by title if omitted)
- `NOTION_RISKS_DB_ID`
- `NOTION_DASHBOARD_PAGE_ID`
- `NOTION_PROJECT_STATUS_PAGE_ID`
- `NOTES_OWNER` (`msp` or `notion`)

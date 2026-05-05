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
- `npm run link:context` (links all Tasks to project context pages/databases such as Dashboard, Project Status, and Risks & Issues)
- `npm run sync:risks` (creates or updates aggregate Risks & Issues records from slipping tasks and links them to Tasks)
- `npm run dashboard:html` (generates `dashboard.html` in the project root)
- `npm run dashboard:pages` (generates `dashboard.html` and copies it to `docs/index.html` for GitHub Pages)
- `npm run build`

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

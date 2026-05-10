# Schedule Intelligence

This is the isolated local schedule analysis app.

It does **not** use Notion.

## What it does

- reads the latest MSP task CSV from `imports/staging/`
- reads supplemental local files:
  - `Project plan with resources and busines.csv`
  - `project team.csv`
  - `Resource table.csv`
- enriches working tasks with likely assignees, business validation owners, and workstreams based on task name matching
- computes a local CPM network from task durations and predecessor links
- generates standalone outputs in `docs/`

## Outputs

- `docs/schedule-intelligence.html`
- `docs/schedule-intelligence-data.json`
- `docs/schedule-intelligence-critical.mmd`

## Run

Use the workspace package script:

`npm run schedule:intelligence`

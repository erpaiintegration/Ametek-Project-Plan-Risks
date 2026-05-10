#!/usr/bin/env python3
"""
Build Resource Action Board - breaks down tasks by person/team with upcoming deadlines.

Reads project plan, RAID issues/actions, and generates:
- Excel workbook with one sheet per resource + dashboard
- JSON for API/programmatic access
- HTML dashboard for browser view
"""

import json
import pandas as pd
import sqlite3
from datetime import datetime, timedelta
from pathlib import Path
import argparse
from urllib.parse import quote


def load_config(config_path):
    """Load configuration from JSON."""
    with open(config_path) as f:
        return json.load(f)


def load_source(source_spec):
    """Load a single source file and apply column mapping."""
    try:
        path = source_spec['path_glob']
        sheet = source_spec.get('sheet_name', 'Sheet1')
        header_row = source_spec.get('header_row', 0)
        
        df = pd.read_excel(path, sheet_name=sheet, header=header_row)
        
        # Apply column mapping if present
        if 'column_map' in source_spec:
            rename_map = {v: k for k, v in source_spec['column_map'].items() if v in df.columns}
            df = df.rename(columns=rename_map)
        
        return df
    except Exception as e:
        print(f"[warn] Could not load {source_spec['name']}: {e}")
        return pd.DataFrame()


def normalize_dates(df):
    """Convert date columns to datetime."""
    for col in ['start_date', 'finish_date', 'baseline_finish', 'actual_start']:
        if col in df.columns:
            df[col] = pd.to_datetime(df[col], errors='coerce')
    return df


def build_resource_tasks(config):
    """
    Load project plan and build action list grouped by resource.
    Returns dict: { resource_name: { 'all_open': [], 'due_14_days': [], 'workstreams': {} } }
    """
    # Load project plan
    plan_src = [s for s in config['sources'] if s['name'] == 'project_plan'][0]
    plan_df = load_source(plan_src)
    plan_df = normalize_dates(plan_df)
    
    # Filter for tasks with resources and open/in-progress status
    plan_df = plan_df[plan_df['owner'].notna() & (plan_df['owner'] != '')]
    
    # Normalize status
    plan_df['status_norm'] = plan_df.get('Status', 'Not Started').fillna('Not Started').str.lower()
    open_statuses = ['future task', 'not started', 'in progress', 'on schedule', 'at risk', 'in queue']
    plan_df = plan_df[plan_df['status_norm'].isin(open_statuses)]
    
    # Calculate due within 14 days
    today = datetime.now().date()
    two_weeks_out = today + timedelta(days=14)
    plan_df['due_soon'] = (plan_df['finish_date'].dt.date >= today) & (plan_df['finish_date'].dt.date <= two_weeks_out)
    
    # Group by resource
    resources = {}
    for owner, group in plan_df.groupby('owner', observed=True):
        # Handle comma-separated resources (e.g., "Person A; Person B")
        for person in str(owner).split(';'):
            person = person.strip()
            if not person or person.lower() == 'nan':
                continue
            
            if person not in resources:
                resources[person] = {
                    'all_open': [],
                    'due_14_days': [],
                    'workstreams': {}
                }
            
            for _, row in group.iterrows():
                task = {
                    'id': row.get('record_id', ''),
                    'name': row.get('task_name', ''),
                    'workstream': row.get('workstream', 'Unassigned'),
                    'status': row.get('Status', 'Not Started'),
                    'priority': row.get('Priority', 'Normal'),
                    'finish': row.get('finish_date'),
                    'percent_complete': row.get('percent_complete', 0),
                    'baseline_finish': row.get('baseline_finish'),
                    'days_due': (row.get('finish_date').date() - today).days if pd.notna(row.get('finish_date')) else 999
                }
                
                # Add to appropriate buckets
                resources[person]['all_open'].append(task)
                
                if row.get('due_soon', False):
                    resources[person]['due_14_days'].append(task)
                
                # Group by workstream
                ws = task['workstream']
                if ws not in resources[person]['workstreams']:
                    resources[person]['workstreams'][ws] = []
                resources[person]['workstreams'][ws].append(task)
    
    # Sort tasks within each resource
    for person in resources:
        resources[person]['due_14_days'] = sorted(
            resources[person]['due_14_days'],
            key=lambda t: (t['days_due'], t['status'].lower())
        )
        resources[person]['all_open'] = sorted(
            resources[person]['all_open'],
            key=lambda t: (t['days_due'] if t['days_due'] < 999 else 999, t['status'].lower())
        )
    
    return resources, plan_df


def load_raid_actions(config):
    """Load RAID actions and filter by resource."""
    try:
        raid_src = [s for s in config['sources'] if s['name'] == 'raid_actions'][0]
        actions_df = load_source(raid_src)
        actions_df = normalize_dates(actions_df)
        
        # Standardize columns if needed
        if 'owner' not in actions_df.columns and 'Owner' in actions_df.columns:
            actions_df.rename(columns={'Owner': 'owner'}, inplace=True)
        
        return actions_df[actions_df['owner'].notna()]
    except Exception as e:
        print(f"[warn] Could not load RAID actions: {e}")
        return pd.DataFrame()


def build_excel_output(resources, plan_df, output_dir):
    """Write resource action board to Excel with multiple sheets."""
    output_path = Path(output_dir) / 'Resource_Action_Board.xlsx'
    
    with pd.ExcelWriter(output_path, engine='openpyxl') as writer:
        # Dashboard sheet
        dashboard_data = []
        for person in sorted(resources.keys()):
            r = resources[person]
            dashboard_data.append({
                'Resource': person,
                'Open Tasks': len(r['all_open']),
                'Due Next 14 Days': len(r['due_14_days']),
                'Workstreams': len(r['workstreams']),
                'At Risk Tasks': sum(1 for t in r['all_open'] if t['status'].lower() in ['at risk', 'behind schedule']),
            })
        
        if dashboard_data:
            dash_df = pd.DataFrame(dashboard_data)
            dash_df.to_excel(writer, sheet_name='Dashboard', index=False)
        
        # One sheet per resource
        for person in sorted(resources.keys()):
            r = resources[person]
            
            if not r['due_14_days']:
                continue
            
            # Build combined task list
            rows = []
            for task in r['due_14_days']:
                rows.append({
                    'Workstream': task['workstream'],
                    'Task': task['name'],
                    'ID': task['id'],
                    'Status': task['status'],
                    'Priority': task['priority'],
                    'Due Date': task['finish'].date() if pd.notna(task['finish']) else '',
                    'Days Until Due': task['days_due'],
                    '% Complete': f"{int(task['percent_complete']*100) if task['percent_complete'] else 0}%",
                })
            
            if rows:
                task_df = pd.DataFrame(rows)
                # Truncate sheet name to 31 chars (Excel limit)
                sheet_name = (person[:28] + '...') if len(person) > 31 else person
                task_df.to_excel(writer, sheet_name=sheet_name, index=False)
    
    return output_path


def build_json_output(resources, output_dir):
    """Write resource action board to JSON."""
    output_path = Path(output_dir) / 'resource_action_board.json'
    
    json_data = {
        'as_of_date': datetime.now().isoformat(),
        'resources': {}
    }
    
    for person in sorted(resources.keys()):
        r = resources[person]
        json_data['resources'][person] = {
            'total_open_tasks': len(r['all_open']),
            'due_next_14_days': len(r['due_14_days']),
            'due_tasks': [
                {
                    'id': t['id'],
                    'name': t['name'],
                    'workstream': t['workstream'],
                    'status': t['status'],
                    'priority': t['priority'],
                    'due_date': t['finish'].isoformat() if pd.notna(t['finish']) else None,
                    'days_due': t['days_due'],
                    'percent_complete': float(t['percent_complete']) if t['percent_complete'] else 0
                }
                for t in r['due_14_days']
            ],
            'by_workstream': {
                ws: len(tasks)
                for ws, tasks in r['workstreams'].items()
            }
        }
    
    with open(output_path, 'w') as f:
        json.dump(json_data, f, indent=2)
    
    return output_path


def build_html_output(resources, output_dir):
    """Write resource action board as interactive HTML dashboard."""
    output_path = Path(output_dir) / 'resource_action_board.html'
    
    html = """<!DOCTYPE html>
<html>
<head>
    <meta charset="utf-8">
    <meta name="viewport" content="width=device-width, initial-scale=1">
    <title>AMETEK SAP S4 - Resource Action Board</title>
    <style>
        body {
            font-family: Segoe UI, Tahoma, Geneva, Verdana, sans-serif;
            margin: 0;
            padding: 20px;
            background: linear-gradient(135deg, #667eea 0%, #764ba2 100%);
            min-height: 100vh;
        }
        .container {
            max-width: 1400px;
            margin: 0 auto;
            background: white;
            border-radius: 8px;
            box-shadow: 0 4px 20px rgba(0,0,0,0.2);
            padding: 30px;
        }
        h1 {
            color: #2c3e50;
            margin-top: 0;
            border-bottom: 3px solid #667eea;
            padding-bottom: 15px;
        }
        .timestamp {
            color: #7f8c8d;
            font-size: 12px;
            margin-bottom: 20px;
        }
        .resource-card {
            background: #f8f9fa;
            border-left: 4px solid #667eea;
            margin: 20px 0;
            padding: 20px;
            border-radius: 4px;
            transition: box-shadow 0.3s;
        }
        .resource-card:hover {
            box-shadow: 0 2px 10px rgba(0,0,0,0.1);
        }
        .resource-name {
            font-size: 18px;
            font-weight: bold;
            color: #2c3e50;
            margin: 0 0 10px 0;
        }
        .resource-stats {
            display: flex;
            gap: 30px;
            margin-bottom: 15px;
            flex-wrap: wrap;
        }
        .stat {
            display: flex;
            align-items: center;
            gap: 8px;
        }
        .stat-value {
            font-size: 20px;
            font-weight: bold;
            color: #667eea;
        }
        .stat-label {
            color: #7f8c8d;
            font-size: 12px;
        }
        .task-list {
            margin-top: 15px;
            border-top: 1px solid #e0e0e0;
            padding-top: 15px;
        }
        .task-item {
            display: flex;
            justify-content: space-between;
            align-items: center;
            padding: 10px;
            margin: 5px 0;
            background: white;
            border-radius: 3px;
            border-left: 3px solid #e0e0e0;
        }
        .task-item.due-soon {
            border-left-color: #e74c3c;
            background: #ffe6e6;
        }
        .task-item.due-imminent {
            border-left-color: #c0392b;
            background: #ffcccc;
        }
        .task-name {
            flex: 1;
            font-weight: 500;
            color: #2c3e50;
        }
        .task-meta {
            display: flex;
            gap: 15px;
            margin-left: 20px;
            font-size: 12px;
            color: #7f8c8d;
        }
        .badge {
            display: inline-block;
            padding: 3px 8px;
            border-radius: 3px;
            font-size: 11px;
            font-weight: bold;
        }
        .badge-workstream {
            background: #ecf0f1;
            color: #2c3e50;
        }
        .badge-status-risk {
            background: #e74c3c;
            color: white;
        }
        .badge-status-ontrack {
            background: #27ae60;
            color: white;
        }
        .badge-status-other {
            background: #3498db;
            color: white;
        }
        .due-date {
            color: #c0392b;
            font-weight: bold;
        }
        .no-tasks {
            color: #95a5a6;
            font-style: italic;
            padding: 20px;
            text-align: center;
        }
    </style>
</head>
<body>
    <div class="container">
        <h1>🎯 AMETEK SAP S4 - Resource Action Board</h1>
        <div class="timestamp">Generated: """ + datetime.now().strftime("%Y-%m-%d %H:%M:%S") + """</div>
"""
    
    for person in sorted(resources.keys()):
        r = resources[person]
        
        if not r['due_14_days']:
            continue
        
        status_badge_map = {
            'at risk': 'badge-status-risk',
            'behind schedule': 'badge-status-risk',
            'on schedule': 'badge-status-ontrack',
            'in progress': 'badge-status-ontrack',
        }
        
        html += f"""
        <div class="resource-card">
            <div class="resource-name">👤 {person}</div>
            <div class="resource-stats">
                <div class="stat">
                    <div class="stat-value">{len(r['all_open'])}</div>
                    <div class="stat-label">Open Tasks</div>
                </div>
                <div class="stat">
                    <div class="stat-value">{len(r['due_14_days'])}</div>
                    <div class="stat-label">Due Next 14 Days</div>
                </div>
                <div class="stat">
                    <div class="stat-value">{len(r['workstreams'])}</div>
                    <div class="stat-label">Workstreams</div>
                </div>
            </div>
            <div class="task-list">
"""
        
        for task in sorted(r['due_14_days'], key=lambda t: t['days_due']):
            status_class = status_badge_map.get(task['status'].lower(), 'badge-status-other')
            due_class = 'due-imminent' if task['days_due'] <= 3 else 'due-soon' if task['days_due'] <= 7 else ''
            
            html += f"""
                <div class="task-item {due_class}">
                    <div class="task-name">{task['name']}</div>
                    <div class="task-meta">
                        <span class="badge badge-workstream">{task['workstream']}</span>
                        <span class="badge {status_class}">{task['status']}</span>
                        <span class="due-date">Due: {task['finish'].strftime('%m/%d') if pd.notna(task['finish']) else 'N/A'} ({task['days_due']}d)</span>
                        <span>{int(task['percent_complete']*100) if task['percent_complete'] else 0}% Complete</span>
                    </div>
                </div>
"""
        
        html += """
            </div>
        </div>
"""
    
    html += """
    </div>
</body>
</html>
"""
    
    with open(output_path, 'w', encoding='utf-8') as f:
        f.write(html)
    
    return output_path


def main():
    parser = argparse.ArgumentParser(description='Build resource action board')
    parser.add_argument('--config', required=True, help='Config JSON path')
    parser.add_argument('--output-dir', required=True, help='Output directory')
    args = parser.parse_args()
    
    config = load_config(args.config)
    resources, plan_df = build_resource_tasks(config)
    
    output_dir = Path(args.output_dir)
    output_dir.mkdir(parents=True, exist_ok=True)
    
    # Build outputs
    print(f"Building resource action board for {len(resources)} resources...")
    
    excel_path = build_excel_output(resources, plan_df, output_dir)
    print(f"[ok] Excel: {excel_path}")
    
    json_path = build_json_output(resources, output_dir)
    print(f"[ok] JSON: {json_path}")
    
    html_path = build_html_output(resources, output_dir)
    print(f"[ok] HTML: {html_path}")
    
    # Summary
    print(f"\nSummary:")
    print(f"  Total resources: {len(resources)}")
    total_open = sum(len(r['all_open']) for r in resources.values())
    total_due = sum(len(r['due_14_days']) for r in resources.values())
    print(f"  Total open tasks: {total_open}")
    print(f"  Total due in 14 days: {total_due}")
    
    print(json.dumps({
        'status': 'success',
        'excel': str(excel_path),
        'json': str(json_path),
        'html': str(html_path),
        'resources': len(resources),
        'total_open_tasks': total_open,
        'total_due_14_days': total_due
    }))


if __name__ == '__main__':
    main()

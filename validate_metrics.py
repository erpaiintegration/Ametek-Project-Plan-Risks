#!/usr/bin/env python3
"""
Comprehensive metric validation script for AMETEK SAP S4 Daily Report
Validates all metrics against source data
"""

import openpyxl
from collections import defaultdict
import pandas as pd
from datetime import datetime

# File paths
RAID_LOG = r"C:\Users\jsw73\OneDrive - AMETEK Inc\SAP Implementation Project - Documents\00 - PMO\04 - Project Tracking and Reporting\AMETEK SAP S4 RAID Log.xlsm"
PROJECT_PLAN = r"C:\Users\jsw73\OneDrive - AMETEK Inc\SAP Implementation Project - Documents\00 - PMO\02 - Project Planning\AMETEK SAP S4 Master Project Plan 3ITC May EXCEL.xlsx"
METRICS_FILE = r"C:\Users\jsw73\OneDrive - AMETEK Inc\Ametek SAP S4 Daily Report\metrics.xlsx"

def load_raid_sheet(sheet_name):
    """Load RAID Log sheet as DataFrame"""
    return pd.read_excel(RAID_LOG, sheet_name=sheet_name, header=0)

def load_project_plan():
    """Load Project Plan sheet as DataFrame"""
    return pd.read_excel(PROJECT_PLAN, sheet_name="Project Plan", header=0)

def load_metrics():
    """Load Metrics sheet as DataFrame"""
    return pd.read_excel(METRICS_FILE, sheet_name="Metrics", header=0)

def validate_risks():
    """Validate Risks metrics"""
    print("\n" + "="*80)
    print("RISKS VALIDATION")
    print("="*80)
    
    risks = load_raid_sheet("Risks")
    # Filter valid risks (R-### format)
    risks = risks[risks['Risk ID'].notna() & risks['Risk ID'].astype(str).str.startswith('R-')]
    
    total = len(risks)
    open_count = len(risks[risks['Risk Status'] == 'Open'])
    closed_count = len(risks[risks['Risk Status'] == 'Closed'])
    
    # Age > 90 days calculation
    today = pd.Timestamp.now()
    risks['Created Date'] = pd.to_datetime(risks['Created Date'], errors='coerce')
    risks['Age Days'] = (today - risks['Created Date']).dt.days
    aged_90_open = len(risks[(risks['Risk Status'] == 'Open') & (risks['Age Days'] > 90)])
    
    print(f"Total Risks:          {total}")
    print(f"  - Open:             {open_count}")
    print(f"  - Closed:           {closed_count}")
    print(f"  - Open > 90 Days:   {aged_90_open}")
    
    expected = {'Total Risks': 91, 'Open': 14, 'Closed': 34, 'Open > 90 Days': 0}
    actual = {'Total Risks': total, 'Open': open_count, 'Closed': closed_count, 'Open > 90 Days': aged_90_open}
    
    print("\nValidation:")
    for key in expected:
        match = "✓" if expected[key] == actual[key] else "✗"
        print(f"  {match} {key}: Expected {expected[key]}, Got {actual[key]}")
    
    return actual, expected

def validate_issues():
    """Validate Issues metrics"""
    print("\n" + "="*80)
    print("ISSUES VALIDATION")
    print("="*80)
    
    issues = load_raid_sheet("Issues")
    issues = issues[issues['Issue ID'].notna() & issues['Issue ID'].astype(str).str.startswith('ISS-')]
    
    total = len(issues)
    open_escalated = len(issues[issues['Issue Status'].isin(['Open', 'Escalated'])])
    escalated = len(issues[issues['Issue Status'] == 'Escalated'])
    closed = len(issues[issues['Issue Status'] == 'Closed'])
    
    # Age > 30 days
    today = pd.Timestamp.now()
    issues['Created Date'] = pd.to_datetime(issues['Created Date'], errors='coerce')
    issues['Age Days'] = (today - issues['Created Date']).dt.days
    aged_30_open = len(issues[(issues['Issue Status'].isin(['Open', 'Escalated'])) & (issues['Age Days'] > 30)])
    
    print(f"Total Issues:         {total}")
    print(f"  - Open/Escalated:   {open_escalated}")
    print(f"  - Escalated:        {escalated}")
    print(f"  - Closed:           {closed}")
    print(f"  - Open > 30 Days:   {aged_30_open}")
    
    expected = {'Total Issues': 32, 'Open/Escalated': 23, 'Escalated': 3, 'Closed': 1, 'Open > 30 Days': 21}
    actual = {'Total Issues': total, 'Open/Escalated': open_escalated, 'Escalated': escalated, 'Closed': closed, 'Open > 30 Days': aged_30_open}
    
    print("\nValidation:")
    for key in expected:
        match = "✓" if expected[key] == actual[key] else "✗"
        print(f"  {match} {key}: Expected {expected[key]}, Got {actual[key]}")
    
    return actual, expected

def validate_decisions():
    """Validate Decisions metrics"""
    print("\n" + "="*80)
    print("DECISIONS VALIDATION")
    print("="*80)
    
    decisions = load_raid_sheet("Decisions")
    decisions = decisions[decisions['Decision ID'].notna() & decisions['Decision ID'].astype(str).str.startswith('D-')]
    
    total = len(decisions)
    open_count = len(decisions[decisions['Decision Status'] == 'Open'])
    closed = len(decisions[decisions['Decision Status'] == 'Closed'])
    pending = len(decisions[decisions['Decision Status'] == 'Pending'])
    
    # Age > 30 days
    today = pd.Timestamp.now()
    decisions['Created Date'] = pd.to_datetime(decisions['Created Date'], errors='coerce')
    decisions['Age Days'] = (today - decisions['Created Date']).dt.days
    aged_30_open = len(decisions[(decisions['Decision Status'] == 'Open') & (decisions['Age Days'] > 30)])
    
    print(f"Total Decisions:      {total}")
    print(f"  - Open:             {open_count}")
    print(f"  - Closed:           {closed}")
    print(f"  - Pending:          {pending}")
    print(f"  - Open > 30 Days:   {aged_30_open}")
    
    expected = {'Total Decisions': 132, 'Open': 12, 'Closed': 114, 'Pending': 0, 'Open > 30 Days': 8}
    actual = {'Total Decisions': total, 'Open': open_count, 'Closed': closed, 'Pending': pending, 'Open > 30 Days': aged_30_open}
    
    print("\nValidation:")
    for key in expected:
        match = "✓" if expected[key] == actual[key] else "✗"
        print(f"  {match} {key}: Expected {expected[key]}, Got {actual[key]}")
    
    return actual, expected

def validate_actions():
    """Validate Actions metrics"""
    print("\n" + "="*80)
    print("ACTIONS VALIDATION")
    print("="*80)
    
    actions = load_raid_sheet("Actions")
    actions = actions[actions['Action Item ID'].notna() & actions['Action Item ID'].astype(str).str.startswith('A-')]
    
    total = len(actions)
    active = len(actions[actions['Action Item Status'].isin(['Open', 'In Progress'])])
    in_progress = len(actions[actions['Action Item Status'] == 'In Progress'])
    closed = len(actions[actions['Action Item Status'] == 'Closed'])
    
    # Overdue: Target Date < Today AND Status not Complete
    today_date = pd.Timestamp.now().date()
    actions['Target Date'] = pd.to_datetime(actions['Target Date'], errors='coerce').dt.date
    overdue = len(actions[(actions['Target Date'] < today_date) & (actions['Action Item Status'] != 'Closed')])
    
    # Avg days overdue
    actions['Days Overdue'] = (today_date - actions['Target Date']).dt.days
    overdue_records = actions[(actions['Target Date'] < today_date) & (actions['Action Item Status'] != 'Closed')]
    avg_days_overdue = round(overdue_records['Days Overdue'].mean(), 1) if len(overdue_records) > 0 else 0
    
    print(f"Total Actions:        {total}")
    print(f"  - Active (O/IP):    {active}")
    print(f"  - In Progress:      {in_progress}")
    print(f"  - Closed:           {closed}")
    print(f"  - Overdue:          {overdue}")
    print(f"  - Avg Days Overdue: {avg_days_overdue}")
    
    expected = {'Total Actions': 205, 'Active': 69, 'In Progress': 15, 'Closed': 126, 'Overdue': 12, 'Avg Days Overdue': 40.6}
    actual = {'Total Actions': total, 'Active': active, 'In Progress': in_progress, 'Closed': closed, 'Overdue': overdue, 'Avg Days Overdue': avg_days_overdue}
    
    print("\nValidation:")
    for key in expected:
        match = "✓" if expected[key] == actual[key] else "✗"
        print(f"  {match} {key}: Expected {expected[key]}, Got {actual[key]}")
    
    return actual, expected

def validate_project_plan():
    """Validate Project Plan metrics"""
    print("\n" + "="*80)
    print("PROJECT PLAN VALIDATION")
    print("="*80)
    
    plan = load_project_plan()
    # Filter to valid tasks (non-summary)
    plan = plan[(plan['Task Name'].notna()) & (plan['Task Name'] != '')]
    
    total = len(plan)
    complete = len(plan[plan['Status'] == 'Complete'])
    on_schedule = len(plan[plan['Status'] == 'On Schedule'])
    late = len(plan[plan['Status'] == 'Late'])
    future = len(plan[plan['Status'] == 'Future Task'])
    
    # Past due: Finish Date < Today AND Status != Complete
    today = pd.Timestamp.now()
    plan['Finish'] = pd.to_datetime(plan['Finish'], errors='coerce')
    past_due = len(plan[(plan['Finish'] < today) & (plan['Status'] != 'Complete')])
    
    # Milestones
    milestones = len(plan[plan['Milestone'].isin([True, 'Yes', 1])])
    
    # Avg % complete
    plan['% Complete'] = pd.to_numeric(plan['% Complete'], errors='coerce')
    avg_pct = round(plan['% Complete'].mean() * 100, 1)
    
    # Avg schedule slip
    plan['Baseline Finish'] = pd.to_datetime(plan['Baseline Finish'], errors='coerce')
    plan['Finish'] = pd.to_datetime(plan['Finish'], errors='coerce')
    plan['Schedule Slip (Days)'] = (plan['Finish'] - plan['Baseline Finish']).dt.days
    plan_with_slip = plan[plan['Schedule Slip (Days)'].notna()]
    avg_slip = round(plan_with_slip['Schedule Slip (Days)'].mean(), 1) if len(plan_with_slip) > 0 else 0
    
    print(f"Total Tasks:          {total}")
    print(f"  - Complete:         {complete}")
    print(f"  - On Schedule:      {on_schedule}")
    print(f"  - Late:             {late}")
    print(f"  - Future Task:      {future}")
    print(f"  - Past Due:         {past_due}")
    print(f"  - Milestones:       {milestones}")
    print(f"  - Avg % Complete:   {avg_pct}%")
    print(f"  - Avg Schedule Slip:{avg_slip} days")
    
    expected = {'Total Tasks': 2219, 'Complete': 489, 'On Schedule': 3, 'Late': 1, 'Future Task': 1726, 
                'Past Due': 13, 'Milestones': 37, 'Avg % Complete': 22.5, 'Avg Schedule Slip': 82.8}
    actual = {'Total Tasks': total, 'Complete': complete, 'On Schedule': on_schedule, 'Late': late, 
              'Future Task': future, 'Past Due': past_due, 'Milestones': milestones, 
              'Avg % Complete': avg_pct, 'Avg Schedule Slip': avg_slip}
    
    print("\nValidation:")
    for key in expected:
        match = "✓" if expected[key] == actual[key] else "✗"
        print(f"  {match} {key}: Expected {expected[key]}, Got {actual[key]}")
    
    return actual, expected

def main():
    print("\n" + "="*80)
    print("AMETEK SAP S4 METRICS VALIDATION")
    print("="*80)
    print(f"Validation Run: {datetime.now().strftime('%Y-%m-%d %H:%M:%S')}")
    
    results = {}
    
    try:
        results['Risks'] = validate_risks()
        results['Issues'] = validate_issues()
        results['Decisions'] = validate_decisions()
        results['Actions'] = validate_actions()
        results['Project Plan'] = validate_project_plan()
        
        # Summary
        print("\n" + "="*80)
        print("VALIDATION SUMMARY")
        print("="*80)
        
        total_checks = 0
        passed_checks = 0
        
        for category, (actual, expected) in results.items():
            for key in expected:
                total_checks += 1
                if expected[key] == actual.get(key):
                    passed_checks += 1
        
        print(f"Total Checks: {passed_checks}/{total_checks}")
        print(f"Pass Rate: {round(passed_checks/total_checks*100, 1)}%")
        
        if passed_checks == total_checks:
            print("\n✓ ALL METRICS VALIDATED SUCCESSFULLY")
        else:
            print("\n✗ SOME METRICS REQUIRE REVIEW")
            
    except Exception as e:
        print(f"\n✗ ERROR: {str(e)}")
        import traceback
        traceback.print_exc()

if __name__ == "__main__":
    main()

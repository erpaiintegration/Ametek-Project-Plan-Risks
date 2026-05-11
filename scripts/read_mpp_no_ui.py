"""
Read MS Project (.mpp) file without UI interaction using win32com.
Sets DisplayAlerts=False and Visible=False to prevent any dialog boxes or window display.
"""
import win32com.client
import pandas as pd
import os
import sys
from datetime import datetime

mpp_path = r"C:\Users\jsw73\OneDrive - AMETEK Inc\SAP Implementation Project - Documents\00 - PMO\02 - Project Planning\AMETEK SAP S4 Master Project Plan 3ITC May.mpp"

if not os.path.exists(mpp_path):
    print(f"Error: File not found at {mpp_path}")
    sys.exit(1)

try:
    # Initialize MS Project application
    msp = win32com.client.Dispatch("MSProject.Application")
    
    # Disable all UI interactions
    msp.Visible = False
    msp.DisplayAlerts = False
    
    # Open the project file
    print(f"Opening {mpp_path}...")
    msp.FileOpen(mpp_path, True, 0)  # msoTrue for read-only, 0 = pjDoNotSave
    
    proj = msp.ActiveProject
    print(f"Project: {proj.Name}")
    
    # Extract all tasks
    tasks_data = []
    for task in proj.Tasks:
        if task is not None:
            try:
                parent_id = task.OutlineParent.ID if task.OutlineLevel > 1 else None
            except:
                parent_id = None
            
            try:
                start_date = str(task.Start) if task.Start else None
                finish_date = str(task.Finish) if task.Finish else None
            except:
                start_date = None
                finish_date = None
            
            tasks_data.append({
                'ID': task.ID,
                'Name': task.Name,
                'OutlineLevel': task.OutlineLevel,
                'ParentID': parent_id,
                'Start': start_date,
                'Finish': finish_date,
                'PercentComplete': task.PercentComplete,
                'UniqueID': task.UniqueID,
            })
    
    df = pd.DataFrame(tasks_data)
    
    # Close without saving
    msp.FileCloseAll(2)  # 2 = pjDoNotSave
    msp.Quit()
    
    print(f"\nSuccessfully loaded {len(df)} tasks")
    print(f"Outline Levels: {sorted(df['OutlineLevel'].unique())}")
    
    # Show workstreams (Level 2)
    print("\n=== WORKSTREAMS (Level 2) ===")
    ws_df = df[df['OutlineLevel'] == 2][['ID', 'Name']]
    for idx, row in ws_df.iterrows():
        print(f"ID {row['ID']:3d}: {row['Name']}")
    
    # Show sample hierarchy
    print("\n=== SAMPLE HIERARCHY CHAINS ===")
    for ws_id in df[df['OutlineLevel'] == 2]['ID'].head(3).tolist():
        ws_name = df[df['ID'] == ws_id].iloc[0]['Name']
        children_l3 = df[df['ParentID'] == ws_id]
        print(f"\n{ws_name} (ID {ws_id}) - {len(children_l3)} Level 3 children")
        for _, child in children_l3.head(3).iterrows():
            grandchildren = df[df['ParentID'] == child['ID']]
            print(f"  ├─ {child['Name']} (ID {child['ID']}, L{child['OutlineLevel']}) - {len(grandchildren)} children")
            for _, gc in grandchildren.head(2).iterrows():
                print(f"  │  ├─ {gc['Name']} (ID {gc['ID']}, L{gc['OutlineLevel']})")
    
    # Export to CSV for inspection
    csv_path = 'outputs/mpp_tasks_export.csv'
    df.to_csv(csv_path, index=False)
    print(f"\n✓ Exported full task data to {csv_path}")
    
except Exception as e:
    print(f"Error: {e}")
    import traceback
    traceback.print_exc()
    try:
        msp.FileCloseAll(2)
        msp.Quit()
    except:
        pass
    sys.exit(1)

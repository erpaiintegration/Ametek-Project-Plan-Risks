import win32com.client
import pandas as pd
import os

mpp_path = r"C:\Users\jsw73\OneDrive - AMETEK Inc\SAP Implementation Project - Documents\00 - PMO\02 - Project Planning\AMETEK SAP S4 Master Project Plan 3ITC May.mpp"

if not os.path.exists(mpp_path):
    print(f"Error: File not found at {mpp_path}")
    exit(1)

try:
    proj_app = win32com.client.Dispatch("MSProject.Application")
    proj_app.Visible = False
    proj_app.FileOpen(mpp_path)
    proj = proj_app.ActiveProject

    tasks_data = []
    for task in proj.Tasks:
        if task:
            try:
                parent_name = task.OutlineParent.Name if task.OutlineLevel > 1 else "ROOT"
                parent_id = task.OutlineParent.ID if task.OutlineLevel > 1 else 0
            except:
                parent_name = None
                parent_id = 0
            
            tasks_data.append({
                'ID': task.ID,
                'Name': task.Name,
                'Level': task.OutlineLevel,
                'ParentID': parent_id,
                'ParentName': parent_name,
                'Start': str(task.Start)[:10] if task.Start else None,
                'Finish': str(task.Finish)[:10] if task.Finish else None,
                'PercentComplete': task.PercentComplete,
            })

    proj_app.FileCloseAll(2) # pjDoNotSave
    proj_app.Quit()

    df = pd.DataFrame(tasks_data)

    print("=== TASK HIERARCHY BY OUTLINE LEVEL ===")
    for level in sorted(df['Level'].unique()):
        level_df = df[df['Level'] == level]
        print(f"\nLevel {level}: {len(level_df)} tasks")
        print(level_df[['ID', 'Name', 'Level', 'ParentName']].head(5).to_string())

    print("\n=== WORKSTREAMS (Level 2) ===")
    workstreams = df[df['Level'] == 2][['ID', 'Name']]
    print(workstreams.to_string())

    print("\n=== SAMPLE HIERARCHY CHAINS (ITC, Data, Testing) ===")
    # Look for IDs matching keywords
    search_ids = []
    for term in ['ITC', 'Data', 'Testing']:
        matches = df[(df['Level'] == 2) & (df['Name'].str.contains(term, case=False, na=False))]
        if not matches.empty:
            search_ids.extend(matches['ID'].tolist())
    
    # Also look at Level 3 if Level 2 is too broad
    if not search_ids:
         for term in ['ITC', 'Data', 'Testing']:
            matches = df[(df['Level'] == 3) & (df['Name'].str.contains(term, case=False, na=False))]
            if not matches.empty:
                search_ids.extend(matches['ID'].tolist())

    for parent_id in list(set(search_ids))[:10]:
        parent = df[df['ID'] == parent_id]
        if not parent.empty:
            parent_name = parent.iloc[0]['Name']
            children_l3 = df[df['ParentID'] == parent_id]
            print(f"\n{parent_name} (ID {parent_id}) [Level {parent.iloc[0]['Level']}]")
            for _, child in children_l3.head(3).iterrows():
                grandchildren = df[df['ParentID'] == child['ID']]
                print(f"  └─ {child['Name']} (ID {child['ID']}) - {len(grandchildren)} children")
                for _, grandchild in grandchildren.head(2).iterrows():
                    print(f"     └─ {grandchild['Name']} (ID {grandchild['ID']})")

except Exception as e:
    print(f"Error: {e}")
    try:
        proj_app.Quit()
    except:
        pass

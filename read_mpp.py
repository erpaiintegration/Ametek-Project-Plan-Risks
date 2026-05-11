import win32com.client
import pandas as pd
import os

mpp_path = r"C:\Users\jsw73\OneDrive - AMETEK Inc\SAP Implementation Project - Documents\00 - PMO\02 - Project Planning\AMETEK SAP S4 Master Project Plan 3ITC May.mpp"

if not os.path.exists(mpp_path):
    print(f"Error: File not found at {mpp_path}")
    exit(1)

try:
    mpp = win32com.client.Dispatch("MSProject.Application")
    mpp.Visible = False
    mpp.FileOpen(mpp_path)
    proj = mpp.ActiveProject
    
    tasks_data = []
    # Project indices are 1-based usually
    for task in proj.Tasks:
        if task is not None:
            tasks_data.append({
                'ID': task.ID,
                'Name': task.Name,
                'OutlineLevel': task.OutlineLevel,
                'Start': str(task.Start),
                'Finish': str(task.Finish),
                'PercentComplete': task.PercentComplete,
            })
    
    df = pd.DataFrame(tasks_data)
    print("Tasks loaded successfully")
    print(f"Total tasks: {len(df)}")
    if not df.empty:
        print(f"Outline levels: {sorted(df['OutlineLevel'].unique())}")
        print("\nFirst 30 tasks by ID:")
        print(df.head(30).to_string())
    
    mpp.FileCloseAll(2) # pjDoNotSave
    mpp.Quit()

except Exception as e:
    print(f"Error: {e}")
    try:
        mpp.Quit()
    except:
        pass

import pandas as pd

df = pd.read_csv("outputs/mpp_tasks_export.csv")

print(f"Total tasks: {len(df)}")
print("\nOutline level distribution:")
print(df["OutlineLevel"].value_counts().sort_index())

workstreams = {2: "System Readiness", 259: "Data Readiness", 2151: "Solution Readiness", 2204: "Business Readiness"}

for ws_id, ws_name in workstreams.items():
    ws_task = df[df["ID"] == ws_id]
    if not ws_task.empty:
        print("\n" + "="*60)
        print(f"{ws_name} (ID {ws_id})")
        print("="*60)
        
        def get_descendants(task_id, level=0, max_depth=3):
            children = df[df["ParentID"] == task_id]
            for _, child in children.iterrows():
                indent = "  " * level
                name = str(child["Name"])[:60]
                print(f"{indent}L{int(child['OutlineLevel'])}: {name}")
                if level < max_depth - 1:
                    get_descendants(child["ID"], level + 1, max_depth)
        
        get_descendants(ws_id)

print("\n" + "="*60)
print("TASK NAME KEYWORDS ANALYSIS")
print("="*60)

keywords = ["ITC", "Test", "Data", "Master", "Smoke", "Ready"]
for kw in keywords:
    matching = df[df["Name"].str.contains(kw, case=False, na=False)]
    print(f"\n'{kw}': {len(matching)} tasks")
    print(matching[["ID", "OutlineLevel", "Name"]].head(3).to_string())

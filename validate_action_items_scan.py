import openpyxl
import pandas as pd

metrics_path = r"C:\Users\jsw73\OneDrive - AMETEK Inc\Ametek SAP S4 Daily Report\Ametek PMO Metrics.xlsx"
huddle_path = r"C:\Users\jsw73\OneDrive - AMETEK Inc\Ametek SAP S4 Daily Report\Ametek SAP S4 PMO Huddle Report.xlsx"

detail = pd.read_excel(metrics_path, sheet_name='PQ_Action_Items_Detail')
print('DETAIL_ROWS', len(detail))
print('SECTION_COUNTS', detail['Section'].value_counts(dropna=False).to_dict())

wb = openpyxl.load_workbook(huddle_path, data_only=True, read_only=True)
ws = wb['Action Items']
targets = [
    'IMMEDIATE + NEEDS ATTENTION SOON',
    'IN PROGRESS + 2 WEEKS LOOKAHEAD',
    'COMBINED ACTION ITEMS METRICS (BY WORKSTREAM / ASSIGNED TO / TYPE)'
]
found = {}
for r in range(1, 4000):
    for c in range(1, 30):
        v = str(ws.cell(r, c).value or '').upper()
        for t in targets:
            if t in v and t not in found:
                found[t] = (r, c)
print('FOUND', found)
wb.close()

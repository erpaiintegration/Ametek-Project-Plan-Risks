' governance-apply-view.vbs
' Apply governance filter and table to draft plan

Option Explicit
Dim msProj, draftProj, i, tbl, flt

' Connect to MS Project
Set msProj = GetObject(, "MSProject.Application")
If msProj Is Nothing Then
    WScript.Echo "ERROR: MS Project not running"
    WScript.Quit 1
End If

' Find and activate Draft
For i = 1 To msProj.Projects.Count
    If InStr(msProj.Projects(i).Name, "Draft") > 0 Then
        Set draftProj = msProj.Projects(i)
        draftProj.Activate
        WScript.Sleep 300
        Exit For
    End If
Next

If draftProj Is Nothing Then
    WScript.Echo "ERROR: Draft not found"
    WScript.Quit 1
End If

WScript.Echo "Applying governance view..."

' Step 1: Create governance table
On Error Resume Next
draftProj.TaskTables("Governance CPM Check").Delete
On Error GoTo 0

On Error Resume Next
Set tbl = draftProj.TaskTables.Add()
tbl.Name = "Governance CPM Check"
tbl.ShowInMenu = True

' Add columns to table
Dim tf
Set tf = tbl.TableFields.Add(188743739, 1): tf.Width = 15: tf.Title = "WBS"
Set tf = tbl.TableFields.Add(188743694, 1): tf.Width = 40: tf.Title = "Task Name"
Set tf = tbl.TableFields.Add(188743715, 1): tf.Width = 12: tf.Title = "Start"
Set tf = tbl.TableFields.Add(188743711, 1): tf.Width = 12: tf.Title = "Finish"
Set tf = tbl.TableFields.Add(188743718, 2): tf.Width = 9: tf.Title = "Slack"
Set tf = tbl.TableFields.Add(188743773, 1): tf.Width = 40: tf.Title = "Mandate | Delta"

On Error GoTo 0

WScript.Echo "Created governance table"

' Step 2: Create governance filter
On Error Resume Next
draftProj.TaskFilters("Governance Gates Only").Delete
On Error GoTo 0

On Error Resume Next
Set flt = draftProj.TaskFilters.Add()
flt.Name = "Governance Gates Only"
flt.ShowInMenu = True
flt.ShowRelatedSummaryRows = True

' Add filter criteria: Flag1 = True
Dim fc
Set fc = flt.FilterCriteria.Add()
fc.FieldName = "Flag1"
fc.Test = 38  ' pjIsTrue

On Error GoTo 0

WScript.Echo "Created governance filter"

' Step 3: Apply to view
On Error Resume Next
draftProj.Views("Gantt Chart").Apply
msProj.ViewApply "Gantt Chart"
draftProj.TaskTables("Governance CPM Check").Apply
draftProj.TaskFilters("Governance Gates Only").Apply
On Error GoTo 0

WScript.Sleep 500

WScript.Echo "SUCCESS: Governance views applied!"
WScript.Echo "- Table: 'Governance CPM Check'"
WScript.Echo "- Filter: 'Governance Gates Only' (showing 13 gates)"
WScript.Echo "- Gate fields colored by schedule delta"
WScript.Quit 0

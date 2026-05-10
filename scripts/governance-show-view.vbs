' governance-show-view.vbs
' Force-activates the Governance CPM Check table + Governance Gates Only filter
' Run this after governance-apply-active.vbs to make the view visible

Option Explicit
Dim msProj

Set msProj = GetObject(, "MSProject.Application")
If msProj Is Nothing Then
    WScript.Echo "ERROR: MS Project not running"
    WScript.Quit 1
End If

On Error Resume Next

' 1. Switch to Gantt Chart view
msProj.ViewApply "Gantt Chart"
WScript.Echo "View: Gantt Chart"

WScript.Sleep 300

' 2. Apply the custom table
msProj.TableApply "Governance CPM Check"
If Err.Number <> 0 Then
    WScript.Echo "Table ERROR: " & Err.Description
    Err.Clear
Else
    WScript.Echo "Table: Governance CPM Check  OK"
End If

WScript.Sleep 300

' 3. Apply the custom filter
msProj.FilterApply "Governance Gates Only"
If Err.Number <> 0 Then
    WScript.Echo "Filter ERROR: " & Err.Description
    Err.Clear
Else
    WScript.Echo "Filter: Governance Gates Only  OK"
End If

' 4. Show what tables/filters exist
WScript.Echo ""
WScript.Echo "Available custom tables:"
Dim tbl
For Each tbl In msProj.ActiveProject.TaskTables
    If InStr(tbl.Name, "Governance") > 0 Or InStr(tbl.Name, "Custom") > 0 Then
        WScript.Echo "  -> " & tbl.Name
    End If
Next

WScript.Echo ""
WScript.Echo "Available custom filters:"
Dim flt
For Each flt In msProj.ActiveProject.TaskFilters
    If InStr(flt.Name, "Governance") > 0 Then
        WScript.Echo "  -> " & flt.Name
    End If
Next

On Error GoTo 0
WScript.Echo ""
WScript.Echo "Done. Check MS Project window."

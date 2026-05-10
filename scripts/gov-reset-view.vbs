Option Explicit
Dim msProj
Set msProj = GetObject(, "MSProject.Application")
If msProj Is Nothing Then WScript.Echo "ERROR" : WScript.Quit 1

WScript.Echo "Resetting to clean view state..."
On Error Resume Next
msProj.ViewApply "Gantt Chart"
WScript.Sleep 200
msProj.TableApply "Entry"
WScript.Sleep 200
msProj.FilterApply "All Tasks"
WScript.Sleep 200
WScript.Echo "Errors so far: " & Err.Number & " " & Err.Description
Err.Clear

' Try to find and list all filters
Dim flt, tbl
WScript.Echo ""
WScript.Echo "Current task filters:"
For Each flt In msProj.ActiveProject.TaskFilters
    WScript.Echo "  " & flt.Name
Next
WScript.Echo ""
WScript.Echo "Current task tables:"
For Each tbl In msProj.ActiveProject.TaskTables
    WScript.Echo "  " & tbl.Name
Next
On Error GoTo 0

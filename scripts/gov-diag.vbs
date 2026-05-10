Option Explicit
Dim msProj
Set msProj = GetObject(, "MSProject.Application")

' List views
On Error Resume Next
Dim v
WScript.Echo "Views:"
For Each v In msProj.Views
    If InStr(v.Name, "Govern") > 0 Then WScript.Echo "  GOVERNANCE: " & v.Name
Next
Err.Clear

' Check global tables
WScript.Echo ""
WScript.Echo "Global tables with Governance:"
Dim gt
For Each gt In msProj.GlobalTaskTables
    If InStr(gt.Name, "Govern") > 0 Then WScript.Echo "  " & gt.Name
Next
Err.Clear

' Check global filters
WScript.Echo ""
WScript.Echo "Global filters with Governance:"
Dim gf
For Each gf In msProj.GlobalTaskFilters
    If InStr(gf.Name, "Govern") > 0 Then WScript.Echo "  " & gf.Name
Next
Err.Clear
On Error GoTo 0
WScript.Echo "Done."

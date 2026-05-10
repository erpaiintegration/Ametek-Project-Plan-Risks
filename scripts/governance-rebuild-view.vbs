Option Explicit
Dim msProj, activeProj

Set msProj = GetObject(, "MSProject.Application")
If msProj Is Nothing Then
    WScript.Echo "ERROR: MS Project not running"
    WScript.Quit 1
End If
Set activeProj = msProj.ActiveProject
If activeProj Is Nothing Then
    WScript.Echo "ERROR: No project open"
    WScript.Quit 1
End If
WScript.Echo "Project: " & Left(activeProj.Name, 60)

' Use TableEditEx - Application-level table creation
' pjTaskWBS=188743739, pjTaskName=188743694, pjTaskStart=188743715
' pjTaskFinish=188743711, pjTaskTotalSlack=188743718, pjTaskText20=188743773

WScript.Echo "[1] Creating table via TableEditEx..."
On Error Resume Next

msProj.TableEditEx _
    "Governance CPM Check", True, "", _
    "WBS", "WBS", 15, 1, 1, _
    "Name", "Task Name", 40, 1, 1, _
    "Start", "Start", 12, 1, 1, _
    "Finish", "Finish", 12, 1, 1, _
    "Total Slack", "Slack", 9, 1, 1, _
    "Text20", "Mandate|Delta", 45, 1, 1

If Err.Number <> 0 Then
    WScript.Echo "FAIL TableEditEx: " & Err.Number & " - " & Err.Description
    Err.Clear
Else
    WScript.Echo "    Table created OK"
End If
On Error GoTo 0

WScript.Sleep 300

WScript.Echo "[2] Creating filter via FilterEdit..."
On Error Resume Next

msProj.FilterEdit "Governance Gates Only", True, True, "Flag1", "equals", "Yes"

If Err.Number <> 0 Then
    WScript.Echo "FAIL FilterEdit: " & Err.Number & " - " & Err.Description
    Err.Clear
Else
    WScript.Echo "    Filter created OK"
End If
On Error GoTo 0

WScript.Sleep 300

WScript.Echo "[3] Applying view..."
On Error Resume Next
msProj.ViewApply "Gantt Chart"
WScript.Sleep 300
msProj.TableApply "Governance CPM Check"
If Err.Number <> 0 Then
    WScript.Echo "FAIL TableApply: " & Err.Number & " - " & Err.Description
    Err.Clear
Else
    WScript.Echo "    TableApply OK"
End If
msProj.FilterApply "Governance Gates Only"
If Err.Number <> 0 Then
    WScript.Echo "FAIL FilterApply: " & Err.Number & " - " & Err.Description
    Err.Clear
Else
    WScript.Echo "    FilterApply OK"
End If
On Error GoTo 0
WScript.Echo "DONE."

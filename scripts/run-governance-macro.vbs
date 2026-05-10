' run-governance-macro.vbs
Option Explicit

Dim msProj, i

On Error Resume Next
Set msProj = GetObject(, "MSProject.Application")
On Error GoTo 0

If msProj Is Nothing Then
    WScript.Echo "ERROR: Microsoft Project is not running."
    WScript.Quit 1
End If

' Make sure app is ready
For i = 1 To 40
    On Error Resume Next
    msProj.Visible = True
    If Err.Number = 0 Then
        On Error GoTo 0
        Exit For
    End If
    Err.Clear
    On Error GoTo 0
    WScript.Sleep 250
Next

' Try running the setup macro from imported module
For i = 1 To 40
    On Error Resume Next
    msProj.RunMacro "SetupGovernanceViews"
    If Err.Number = 0 Then
        On Error GoTo 0
        WScript.Echo "OK: SetupGovernanceViews completed."
        WScript.Quit 0
    End If
    Err.Clear
    On Error GoTo 0
    WScript.Sleep 250
Next

WScript.Echo "ERROR: Could not run 'SetupGovernanceViews'. Import scripts\\governance-views.bas into VBA and run SetupGovernanceViews manually (F5)."
WScript.Quit 2

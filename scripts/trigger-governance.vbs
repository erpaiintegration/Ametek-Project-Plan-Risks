' trigger-governance.vbs
' Runs pre-imported governance macro with robust COM retry
' Requires: governance-simple.bas already imported into MS Project VBA

Option Explicit
Dim msProj, k, success

Function TryGetApp()
    Dim app
    On Error Resume Next
    For k = 1 To 10
        Set app = GetObject(, "MSProject.Application")
        If Err.Number = 0 Then
            Set TryGetApp = app
            Exit Function
        End If
        Err.Clear
        WScript.Sleep 500
    Next
    Set TryGetApp = Nothing
End Function

Set msProj = TryGetApp()

If msProj Is Nothing Then
    WScript.Echo "ERROR: MS Project not found"
    WScript.Quit 1
End If

If msProj.ActiveProject Is Nothing Then
    WScript.Echo "ERROR: No project open. Open draft plan first."
    WScript.Quit 1
End If

WScript.Echo "Running governance views on: " & msProj.ActiveProject.Name

success = 0
For k = 1 To 30
    On Error Resume Next
    msProj.RunMacro "GovernanceSimple.RunGovernanceViews"
    If Err.Number = 0 Then
        success = 1
        On Error GoTo 0
        Exit For
    End If
    WScript.Echo "Retry " & k & "/30 (COM busy)..."
    Err.Clear
    On Error GoTo 0
    WScript.Sleep 1000
Next

If success = 1 Then
    WScript.Echo "OK: Governance views applied!"
    WScript.Quit 0
Else
    WScript.Echo "ERROR: Macro run failed"
    WScript.Quit 2
End If

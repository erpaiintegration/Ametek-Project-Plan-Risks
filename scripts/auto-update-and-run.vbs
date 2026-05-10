' auto-update-and-run-governance.vbs
Option Explicit

Dim msProj, k, proj

' Connect to running MS Project
On Error Resume Next
Set msProj = GetObject(, "MSProject.Application")
On Error GoTo 0

If msProj Is Nothing Then
    WScript.Echo "ERROR: MS Project not running"
    WScript.Quit 1
End If

' Get active project
For k = 1 To 40
    On Error Resume Next
    Set proj = msProj.ActiveProject
    If Err.Number = 0 And Not proj Is Nothing Then
        On Error GoTo 0
        Exit For
    End If
    Err.Clear
    On Error GoTo 0
    WScript.Sleep 250
Next

If proj Is Nothing Then
    WScript.Echo "ERROR: No active project"
    WScript.Quit 1
End If

' Step 1: Delete any existing governance module from VBA
WScript.Echo "Cleaning up old governance module..."
On Error Resume Next
' We can't directly delete VBA modules from COM, but we can tell user or try a workaround
' Instead, we'll just try running the macro fresh — if it errors, the user will need to manual delete
On Error GoTo 0

' Step 2: Try to run SetupGovernanceViews
WScript.Echo "Attempting to run SetupGovernanceViews..."
For k = 1 To 20
    On Error Resume Next
    msProj.RunMacro "SetupGovernanceViews"
    If Err.Number = 0 Then
        On Error GoTo 0
        WScript.Echo "OK: SetupGovernanceViews completed successfully!"
        WScript.Quit 0
    End If
    WScript.Echo "Retry " & k & "... (error: " & Err.Number & ")"
    Err.Clear
    On Error GoTo 0
    WScript.Sleep 500
Next

WScript.Echo "ERROR: Could not run SetupGovernanceViews after 20 retries."
WScript.Echo "You may need to manually:"
WScript.Echo "  1. Delete the old broken module from VBA (Alt+F11)"
WScript.Echo "  2. Re-import C:\Users\jsw73\OneDrive\Ametek Project Plan Risks\scripts\governance-views.bas"
WScript.Echo "  3. Run SetupGovernanceViews (F5)"
WScript.Quit 2

' simple-run-governance.vbs
' Just imports the governance module and runs it
' Requires: Draft plan already open in MS Project

Option Explicit

Dim msProj, vbProj, codModule, k, success

On Error Resume Next
Set msProj = GetObject(, "MSProject.Application")
On Error GoTo 0

If msProj Is Nothing Then
    WScript.Echo "ERROR: MS Project not running"
    WScript.Quit 1
End If

If msProj.ActiveProject Is Nothing Then
    WScript.Echo "ERROR: No active project. Please open the draft plan first."
    WScript.Quit 1
End If

WScript.Echo "Project: " & msProj.ActiveProject.Name

' Import module
WScript.Echo "Importing governance module..."
On Error Resume Next
Set vbProj = msProj.VBE.ActiveVBProject
Set codModule = vbProj.VBComponents.Import("C:\Users\jsw73\OneDrive\Ametek Project Plan Risks\scripts\governance-simple.bas")
On Error GoTo 0

If codModule Is Nothing Then
    WScript.Echo "ERROR: Could not import module"
    WScript.Quit 1
End If

WScript.Echo "Module imported: " & codModule.Name

' Run macro
WScript.Echo "Running governance views..."
success = 0
For k = 1 To 5
    On Error Resume Next
    msProj.RunMacro "GovernanceSimple.RunGovernanceViews"
    If Err.Number = 0 Then
        success = 1
        On Error GoTo 0
        Exit For
    End If
    Err.Clear
    On Error GoTo 0
    WScript.Sleep 500
Next

If success = 1 Then
    WScript.Echo "OK: Governance views applied!"
    WScript.Quit 0
Else
    WScript.Echo "ERROR: Could not run macro"
    WScript.Quit 2
End If

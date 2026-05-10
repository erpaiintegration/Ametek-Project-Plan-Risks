' governance-apply-direct.vbs
' Direct import and run - VBScript version
' Run from command line: cscript governance-apply-direct.vbs

Option Explicit

Dim msProj, vbProj, codModule, k, basPath, success

basPath = "C:\Users\jsw73\OneDrive\Ametek Project Plan Risks\scripts\governance-simple.bas"

WScript.Echo "Connecting to MS Project..."
On Error Resume Next
Set msProj = GetObject(, "MSProject.Application")
On Error GoTo 0

If msProj Is Nothing Then
    WScript.Echo "ERROR: MS Project not running"
    WScript.Quit 1
End If

WScript.Echo "OK"

' Try to find the DRAFT file in open projects
Dim proj, i, foundDraft
Set proj = Nothing
foundDraft = 0

WScript.Echo "Looking for Draft file..."
On Error Resume Next

For i = 1 To msProj.Projects.Count
    If Err.Number = 0 Then
        Set proj = msProj.Projects(i)
        WScript.Echo "Found: " & proj.Name
        
        If InStr(proj.Name, "Draft") > 0 Then
            foundDraft = 1
            WScript.Echo "*** FOUND DRAFT ***"
            Exit For
        End If
    End If
Next

On Error GoTo 0

If foundDraft = 0 Or proj Is Nothing Then
    WScript.Echo "ERROR: Could not find Draft file. Please ensure 'AMETEK SAP S4 Master Project Plan May Full Replan Draft.mpp' is open."
    WScript.Quit 1
End If

WScript.Echo "Activating: " & proj.Name
On Error Resume Next
proj.Activate
On Error GoTo 0

WScript.Sleep 1000

' Import module
WScript.Echo "Importing governance module..."
On Error Resume Next
Set vbProj = msProj.VBE.ActiveVBProject
Set codModule = vbProj.VBComponents.Import(basPath)
On Error GoTo 0

If codModule Is Nothing Then
    WScript.Echo "ERROR: Could not import module"
    WScript.Quit 1
End If

Dim moduleName
moduleName = codModule.Name
WScript.Echo "Imported as: " & moduleName
WScript.Sleep 1000

' Run macro
WScript.Echo "Running governance views..."
success = 0
Dim errNum, errDesc
For k = 1 To 10
    On Error Resume Next
    
    ' Try with module prefix first
    msProj.RunMacro moduleName & ".RunGovernanceViews"
    errNum = Err.Number
    errDesc = Err.Description
    
    ' If that fails, try without prefix
    If errNum <> 0 Then
        Err.Clear
        msProj.RunMacro "RunGovernanceViews"
        errNum = Err.Number
        errDesc = Err.Description
    End If
    
    On Error GoTo 0
    
    If errNum = 0 Then
        success = 1
        WScript.Echo "SUCCESS!"
        Exit For
    End If
    
    WScript.Echo "Attempt " & k & ": Error " & errNum & " - " & errDesc
    WScript.Sleep 500
Next

If success = 1 Then
    WScript.Echo "SUCCESS: Governance views applied!"
    WScript.Quit 0
Else
    WScript.Echo "ERROR: Could not run macro"
    WScript.Quit 2
End If

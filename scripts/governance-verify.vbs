' governance-verify.vbs
' Check and verify all governance updates are in place

Option Explicit
Dim msProj, draftProj, i, task, flagCount, colorCount, fieldCount

Set msProj = GetObject(, "MSProject.Application")
If msProj Is Nothing Then
    WScript.Echo "ERROR: MS Project not running"
    WScript.Quit 1
End If

' Find draft
For i = 1 To msProj.Projects.Count
    If InStr(msProj.Projects(i).Name, "Draft") > 0 Then
        Set draftProj = msProj.Projects(i)
        Exit For
    End If
Next

If draftProj Is Nothing Then
    WScript.Echo "ERROR: Draft not found"
    WScript.Quit 1
End If

WScript.Echo "GOVERNANCE VERIFICATION REPORT"
WScript.Echo "==============================="
WScript.Echo "Project: " & Left(draftProj.Name, 60)
WScript.Echo ""

' === CHECK FLAGS ===
flagCount = 0
colorCount = 0
fieldCount = 0

On Error Resume Next
For Each task In draftProj.Tasks
    If Not task Is Nothing Then
        If task.Flag1 = True Then
            flagCount = flagCount + 1
            
            ' Check if field has data
            If Len(task.Text20) > 0 Then
                fieldCount = fieldCount + 1
            End If
            
            ' Check if colored (font color not default)
            If task.Font.Color <> 0 Then
                colorCount = colorCount + 1
            End If
        End If
    End If
Next
On Error GoTo 0

WScript.Echo "[GATES STAMPED]"
WScript.Echo "  Flag1 (marked): " & flagCount
WScript.Echo "  Text20 (fields): " & fieldCount
WScript.Echo "  Font.Color (colored): " & colorCount
WScript.Echo ""

' === CHECK TABLE ===
Dim tblExists, tblActive
tblExists = 0
tblActive = 0

On Error Resume Next
If Not draftProj.TaskTables("Governance CPM Check") Is Nothing Then
    tblExists = 1
End If
On Error GoTo 0

WScript.Echo "[TABLE]"
If tblExists = 1 Then
    WScript.Echo "  Status: Governance CPM Check EXISTS"
Else
    WScript.Echo "  Status: Governance CPM Check NOT FOUND (re-run apply script)"
End If
WScript.Echo ""

' === CHECK FILTER ===
Dim fltExists
fltExists = 0

On Error Resume Next
If Not draftProj.TaskFilters("Governance Gates Only") Is Nothing Then
    fltExists = 1
End If
On Error GoTo 0

WScript.Echo "[FILTER]"
If fltExists = 1 Then
    WScript.Echo "  Status: Governance Gates Only EXISTS"
Else
    WScript.Echo "  Status: Governance Gates Only NOT FOUND (re-run apply script)"
End If
WScript.Echo ""

' === SUMMARY ===
WScript.Echo "[SUMMARY]"
If flagCount >= 13 And fieldCount >= 13 And colorCount >= 13 Then
    WScript.Echo "  Status: COMPLETE ✓"
    WScript.Echo "  All 13 gates stamped, colored, and marked with deltas"
    WScript.Quit 0
Else
    WScript.Echo "  Status: INCOMPLETE"
    WScript.Echo "  Gates: " & flagCount & "/13"
    WScript.Echo "  Fields: " & fieldCount & "/13"
    WScript.Echo "  Colors: " & colorCount & "/13"
    WScript.Echo ""
    WScript.Echo "  Recommendation: Re-run governance-apply-all.vbs to complete setup"
    WScript.Quit 1
End If

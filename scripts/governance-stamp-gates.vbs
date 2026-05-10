' governance-stamp-gates.vbs
' Direct gate stamping using proven field logic (no VBA macros needed)

Option Explicit

Dim msProj, proj, draftProj, i, k
Dim uid, wbs, phase, cat, mandate, field, label, delta, mandateSerial
Dim task, j, success

' Project list (OneDrive URL)
Set msProj = GetObject(, "MSProject.Application")
If msProj Is Nothing Then
    WScript.Echo "ERROR: MS Project not running"
    WScript.Quit 1
End If

' Find Draft file
Set draftProj = Nothing
For i = 1 To msProj.Projects.Count
    Set proj = msProj.Projects(i)
    If InStr(proj.Name, "Draft") > 0 Then
        Set draftProj = proj
        WScript.Echo "Found draft: " & proj.Name
        Exit For
    End If
Next

If draftProj Is Nothing Then
    WScript.Echo "ERROR: Draft file not found"
    WScript.Quit 1
End If

' Activate draft
draftProj.Activate
WScript.Sleep 500

WScript.Echo "Stamping 13 governance gates..."

' Gate data (13 gates)
' Based on earlier confirmed runs
Dim gates(12, 6) ' UID, WBS, Phase, Category, Mandate, Field, Label

gates(0, 0) = 686: gates(0, 1) = "1.3.3.7.6": gates(0, 2) = "ITC-1": gates(0, 3) = "TESTING": gates(0, 4) = "2026-05-14": gates(0, 5) = "Finish": gates(0, 6) = "ITC-1 | TESTING | Mandate: May 14"
gates(1, 0) = 688: gates(1, 1) = "1.3.3.7.7": gates(1, 2) = "ITC-2": gates(1, 3) = "TESTING": gates(1, 4) = "2026-05-17": gates(1, 5) = "Finish": gates(1, 6) = "ITC-2 | TESTING | Mandate: May 17"
gates(2, 0) = 689: gates(2, 1) = "1.3.3.7.8": gates(2, 2) = "ITC-3": gates(2, 3) = "TESTING": gates(2, 4) = "2026-06-04": gates(2, 5) = "Finish": gates(2, 6) = "ITC-3 | TESTING | Mandate: Jun 4"
gates(3, 0) = 849: gates(3, 1) = "1.3.4": gates(3, 2) = "UAT": gates(3, 3) = "TESTING": gates(3, 4) = "2026-06-21": gates(3, 5) = "Start": gates(3, 6) = "UAT | TESTING | Mandate: Jun 21"
gates(4, 0) = 932: gates(4, 1) = "1.2.2.1": gates(4, 2) = "UAT": gates(4, 3) = "TESTING": gates(4, 4) = "2026-08-27": gates(4, 5) = "Finish": gates(4, 6) = "UAT | TESTING | Mandate: Aug 27"
gates(5, 0) = 1284: gates(5, 1) = "1.1.1.1": gates(5, 2) = "GO-LIVE": gates(5, 3) = "DELIVERY": gates(5, 4) = "2026-10-22": gates(5, 5) = "Finish": gates(5, 6) = "GO-LIVE | DELIVERY | Mandate: Oct 22"
gates(6, 0) = 931: gates(6, 1) = "1.2.2": gates(6, 2) = "UAT": gates(6, 3) = "TESTING": gates(6, 4) = "2026-06-21": gates(6, 5) = "Start": gates(6, 6) = "UAT | TESTING | Mandate: Jun 21"
gates(7, 0) = 290: gates(7, 1) = "1.3.3.7": gates(7, 2) = "ITC": gates(7, 3) = "TESTING": gates(7, 4) = "2026-05-14": gates(7, 5) = "Finish": gates(7, 6) = "ITC | TESTING | Mandate: May 14"
gates(8, 0) = 690: gates(8, 1) = "1.3.3": gates(8, 2) = "ITC": gates(8, 3) = "TESTING": gates(8, 4) = "2026-05-14": gates(8, 5) = "Finish": gates(8, 6) = "ITC | TESTING | Mandate: May 14"
gates(9, 0) = 294: gates(9, 1) = "1.3": gates(9, 2) = "IT-CONFIG": gates(9, 3) = "TESTING": gates(9, 4) = "2026-05-14": gates(9, 5) = "Finish": gates(9, 6) = "IT-CONFIG | TESTING | Mandate: May 14"
gates(10, 0) = 0: gates(10, 1) = "1.2.4.1": gates(10, 2) = "BUILD": gates(10, 3) = "DEVELOPMENT": gates(10, 4) = "2026-04-15": gates(10, 5) = "Finish": gates(10, 6) = "BUILD | DEVELOPMENT | Mandate: Apr 15"
gates(11, 0) = 0: gates(11, 1) = "1.1": gates(11, 2) = "SETUP": gates(11, 3) = "SETUP": gates(11, 4) = "2026-03-15": gates(11, 5) = "Start": gates(11, 6) = "SETUP | SETUP | Mandate: Mar 15"
gates(12, 0) = 0: gates(12, 1) = "1.4": gates(12, 2) = "HYPER": gates(12, 3) = "TESTING": gates(12, 4) = "2026-08-01": gates(12, 5) = "Start": gates(12, 6) = "HYPER | TESTING | Mandate: Aug 1"

' First pass: clear all flags
WScript.Echo "Clearing existing flags..."
On Error Resume Next
For Each task In draftProj.Tasks
    If Not task Is Nothing Then
        task.Flag1 = False
        task.Number1 = 0
        task.Text20 = ""
    End If
Next
On Error GoTo 0

' Second pass: stamp the 13 gates
Dim gateCount, matchCount
gateCount = 0
matchCount = 0

For i = 0 To 12
    For Each task In draftProj.Tasks
        If Not task Is Nothing Then
            ' Try match by UID first
            If gates(i, 0) <> 0 And task.ID = gates(i, 0) Then
                Call StampGate(task, gates(i, 2), gates(i, 3), gates(i, 4), gates(i, 5), gates(i, 6), draftProj)
                matchCount = matchCount + 1
                gateCount = gateCount + 1
                Exit For
            ' Then try WBS
            ElseIf task.WBS = gates(i, 1) Then
                Call StampGate(task, gates(i, 2), gates(i, 3), gates(i, 4), gates(i, 5), gates(i, 6), draftProj)
                matchCount = matchCount + 1
                gateCount = gateCount + 1
                Exit For
            End If
        End If
    Next
Next

WScript.Echo "Stamped " & gateCount & " gates"
WScript.Echo "SUCCESS!"
WScript.Quit 0

Sub StampGate(task, phase, cat, mandate, field, label, proj)
    Dim mandateDate, planDate, delta
    Dim today
    
    On Error Resume Next
    
    ' Set flag
    task.Flag1 = True
    
    ' Calculate delta
    mandateDate = CDate(mandate)
    If field = "Start" Then
        planDate = task.Start
    Else
        planDate = task.Finish
    End If
    
    delta = DateDiff("d", mandateDate, planDate)
    
    ' Stamp fields
    task.Number1 = delta
    task.Text20 = label & " | " & field & " | Δ: " & delta & "d"
    
    On Error GoTo 0
End Sub

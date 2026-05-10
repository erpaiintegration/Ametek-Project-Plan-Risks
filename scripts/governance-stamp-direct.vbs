' governance-stamp-direct.vbs
' Ultra-simple: just stamp the 13 known gates by ID

Option Explicit
Dim msProj, draftProj, i, task, gated, stamped

' Connect
Set msProj = GetObject(, "MSProject.Application")
If msProj Is Nothing Then
    WScript.Echo "ERROR: MS Project not running"
    WScript.Quit 1
End If

' Find and activate Draft
For i = 1 To msProj.Projects.Count
    If InStr(msProj.Projects(i).Name, "Draft") > 0 Then
        Set draftProj = msProj.Projects(i)
        draftProj.Activate
        WScript.Sleep 300
        Exit For
    End If
Next

If draftProj Is Nothing Then
    WScript.Echo "ERROR: Draft not found"
    WScript.Quit 1
End If

WScript.Echo "Found draft, stamping 13 gates..."

' The 13 known gate IDs (task.ID) with gate name, mandate date, and field
' Format: Array(ID, WBS, GateName, MandateDate, Field)
Dim gates(12, 4)

gates(0, 0) = 686:   gates(0, 1) = "": gates(0, 2) = "ITC-1": gates(0, 3) = "2026-05-14": gates(0, 4) = "Finish"
gates(1, 0) = 688:   gates(1, 1) = "": gates(1, 2) = "ITC-2": gates(1, 3) = "2026-05-17": gates(1, 4) = "Finish"
gates(2, 0) = 689:   gates(2, 1) = "": gates(2, 2) = "ITC-3": gates(2, 3) = "2026-06-04": gates(2, 4) = "Finish"
gates(3, 0) = 849:   gates(3, 1) = "": gates(3, 2) = "UAT": gates(3, 3) = "2026-06-21": gates(3, 4) = "Start"
gates(4, 0) = 932:   gates(4, 1) = "": gates(4, 2) = "UAT": gates(4, 3) = "2026-08-27": gates(4, 4) = "Finish"
gates(5, 0) = 1284:  gates(5, 1) = "": gates(5, 2) = "GO-LIVE": gates(5, 3) = "2026-10-22": gates(5, 4) = "Finish"
gates(6, 0) = 931:   gates(6, 1) = "": gates(6, 2) = "UAT": gates(6, 3) = "2026-06-21": gates(6, 4) = "Start"
gates(7, 0) = 290:   gates(7, 1) = "": gates(7, 2) = "ITC": gates(7, 3) = "2026-05-14": gates(7, 4) = "Finish"
gates(8, 0) = 690:   gates(8, 1) = "": gates(8, 2) = "ITC": gates(8, 3) = "2026-05-14": gates(8, 4) = "Finish"
gates(9, 0) = 294:   gates(9, 1) = "": gates(9, 2) = "IT-CONFIG": gates(9, 3) = "2026-05-14": gates(9, 4) = "Finish"
gates(10, 0) = 0:    gates(10, 1) = "1.2.4.1": gates(10, 2) = "BUILD": gates(10, 3) = "2026-04-15": gates(10, 4) = "Finish"
gates(11, 0) = 0:    gates(11, 1) = "1.1": gates(11, 2) = "SETUP": gates(11, 3) = "2026-03-15": gates(11, 4) = "Start"
gates(12, 0) = 0:    gates(12, 1) = "1.4": gates(12, 2) = "HYPER": gates(12, 3) = "2026-08-01": gates(12, 4) = "Start"

stamped = 0

' Stamp by ID first (0-9)
On Error Resume Next
For i = 0 To 9
    Set task = draftProj.Tasks(gates(i, 0))
    If Not task Is Nothing Then
        Call StampTask(task, gates(i, 2), gates(i, 3), gates(i, 4))
        stamped = stamped + 1
    End If
Next
On Error GoTo 0

' Stamp by WBS (10-12)
On Error Resume Next
For i = 10 To 12
    For Each task In draftProj.Tasks
        If Not task Is Nothing Then
            If task.WBS = gates(i, 1) Then
                Call StampTask(task, gates(i, 2), gates(i, 3), gates(i, 4))
                stamped = stamped + 1
                Exit For
            End If
        End If
    Next
Next
On Error GoTo 0

WScript.Echo "Stamped " & stamped & " gates by direct ID match"
WScript.Echo "SUCCESS"
WScript.Quit 0

Sub StampTask(tsk, gateName, mandate, field)
    Dim mandateDate, planDate, delta
    Dim color
    
    On Error Resume Next
    
    tsk.Flag1 = True
    
    ' Calculate delta
    mandateDate = CDate(mandate)
    If field = "Start" Then
        planDate = tsk.Start
    Else
        planDate = tsk.Finish
    End If
    
    delta = DateDiff("d", mandateDate, planDate)
    tsk.Number1 = delta
    tsk.Text20 = gateName & " | Mandate: " & mandate & " | Delta: " & delta & "d"
    
    ' Color coding
    If tsk.PercentComplete >= 100 Then
        color = 16711680  ' BLUE
    ElseIf delta > 30 Then
        color = 255        ' RED
    ElseIf delta > 0 Then
        color = 22550 + (80 * 256)  ' ORANGE
    ElseIf delta >= -4 Then
        color = 41120 + (100 * 256)  ' AMBER
    Else
        color = 5287936    ' GREEN
    End If
    
    tsk.Font.Color = color
    tsk.Font.Bold = True
    
    On Error GoTo 0
End Sub

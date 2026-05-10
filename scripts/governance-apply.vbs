' governance-apply.vbs
' Run: Right-click → Open (or double-click)
' Connects to MS Project, opens the draft plan, stamps governance fields,
' creates table/filter, applies color coding — all via COM automation.

Option Explicit

Dim msProj, proj, filePath
Dim gates(), i, t, found
Dim planDate, mandateDate, delta

Function TrySetVisible(appObj, retries, delayMs)
    Dim k
    TrySetVisible = False
    For k = 1 To retries
        On Error Resume Next
        appObj.Visible = True
        If Err.Number = 0 Then
            On Error GoTo 0
            TrySetVisible = True
            Exit Function
        End If
        Err.Clear
        On Error GoTo 0
        WScript.Sleep delayMs
    Next
End Function

filePath = "C:\Users\jsw73\OneDrive\Ametek Project Plan Risks\AMETEK SAP S4 Master Project Plan May Full Replan Draft.mpp"

' ---- Connect to running MS Project ----
On Error Resume Next
Set msProj = GetObject(, "MSProject.Application")
If Err.Number <> 0 Then
    Err.Clear
    Set msProj = CreateObject("MSProject.Application")
End If
On Error GoTo 0

If msProj Is Nothing Then
    MsgBox "Could not connect to Microsoft Project.", vbCritical, "Error"
    WScript.Quit 1
End If

If Not TrySetVisible(msProj, 40, 500) Then
    MsgBox "Microsoft Project is busy (COM call rejected). Please bring Project to foreground, dismiss open dialogs, and re-run.", vbExclamation, "Project Busy"
    WScript.Quit 1
End If

' ---- Open the file (or use already open) ----
Dim alreadyOpen : alreadyOpen = False
On Error Resume Next
Set proj = msProj.ActiveProject
If Not proj Is Nothing Then
    If InStr(1, proj.FullName, "Draft", 1) > 0 Then alreadyOpen = True
End If
On Error GoTo 0

If Not alreadyOpen Then
    msProj.FileOpen filePath, False
    WScript.Sleep 8000
    
    ' Dismiss any dialogs (Read-Only prompt, etc.)
    Dim shell : Set shell = CreateObject("WScript.Shell")
    shell.AppActivate "Microsoft Project"
    WScript.Sleep 1000
    shell.SendKeys "{ENTER}"
    WScript.Sleep 3000
    shell.SendKeys "{ENTER}"
    WScript.Sleep 3000
    
    Set proj = msProj.ActiveProject
End If

If proj Is Nothing Then
    MsgBox "Could not load project. Please open the file manually in MS Project, then run this script again.", vbCritical, "Error"
    WScript.Quit 1
End If

Dim taskCount : taskCount = 0
On Error Resume Next : taskCount = proj.Tasks.Count : On Error GoTo 0

If taskCount = 0 Then
    WScript.Sleep 5000
    On Error Resume Next : taskCount = proj.Tasks.Count : On Error GoTo 0
End If

MsgBox "Connected to: " & proj.Name & vbCrLf & "Tasks: " & taskCount & vbCrLf & vbCrLf & "Click OK to apply governance views.", vbInformation, "AMETEK Governance"

' ============================================================
' GOVERNANCE GATE DATA
' Array: UID (0=use WBS), WBS, Phase, Cat, MandateDate, Field(S/F), Label
' ============================================================
Dim gUID(12), gWBS(12), gPhase(12), gCat(12), gMandate(12), gField(12), gLabel(12)

gUID(0)=686:  gWBS(0)="":          gPhase(0)="ITC-1":    gCat(0)="TESTING":   gMandate(0)="5/14/2026": gField(0)="F": gLabel(0)="ITC-1 Team Testing"
gUID(1)=688:  gWBS(1)="":          gPhase(1)="ITC-1":    gCat(1)="DEFECT":    gMandate(1)="5/14/2026": gField(1)="F": gLabel(1)="ITC-1 Defect Correction"
gUID(2)=689:  gWBS(2)="":          gPhase(2)="ITC-1":    gCat(2)="GATE":      gMandate(2)="5/14/2026": gField(2)="F": gLabel(2)="ITC-1 Gate Review"
gUID(3)=849:  gWBS(3)="":          gPhase(3)="ITC-2":    gCat(3)="REFRESH":   gMandate(3)="5/17/2026": gField(3)="F": gLabel(3)="QS4 ITC-2 Refresh"
gUID(4)=932:  gWBS(4)="":          gPhase(4)="ITC-2":    gCat(4)="TRR":       gMandate(4)="5/17/2026": gField(4)="F": gLabel(4)="TRR ITC-1 (feeds ITC-2)"
gUID(5)=1284: gWBS(5)="":          gPhase(5)="ITC-3":    gCat(5)="REFRESH":   gMandate(5)="6/4/2026":  gField(5)="F": gLabel(5)="QS5 ITC-3 Refresh"
gUID(6)=931:  gWBS(6)="":          gPhase(6)="ITC-3":    gCat(6)="TRR":       gMandate(6)="6/4/2026":  gField(6)="F": gLabel(6)="TRR ITC-3"
gUID(7)=0:    gWBS(7)="1.3.3.7.6": gPhase(7)="ITC-3":    gCat(7)="KICKOFF":   gMandate(7)="7/12/2026": gField(7)="S": gLabel(7)="ITC-3 Test Kickoff (MSO)"
gUID(8)=884:  gWBS(8)="":          gPhase(8)="UAT":      gCat(8)="REFRESH":   gMandate(8)="6/21/2026": gField(8)="F": gLabel(8)="QS4 UAT Refresh"
gUID(9)=0:    gWBS(9)="1.2.4.1":   gPhase(9)="UAT":      gCat(9)="DATA GATE": gMandate(9)="6/21/2026": gField(9)="S": gLabel(9)="PMO Auth UAT Data Loads"
gUID(10)=290: gWBS(10)="":         gPhase(10)="UAT":     gCat(10)="TESTING":  gMandate(10)="8/27/2026":gField(10)="F":gLabel(10)="UAT with Roles"
gUID(11)=690: gWBS(11)="":         gPhase(11)="UAT":     gCat(11)="TRR":      gMandate(11)="8/27/2026":gField(11)="F":gLabel(11)="TRR UAT"
gUID(12)=294: gWBS(12)="":         gPhase(12)="GO-LIVE": gCat(12)="MILESTONE":gMandate(12)="10/22/2026":gField(12)="F":gLabel(12)="Go-Live Oct 22"

' ============================================================
' STEP 1: Clear existing flags, stamp governance fields
' ============================================================
Dim stamped : stamped = 0
Dim isMatch, statusTxt

For Each t In proj.Tasks
    If Not t Is Nothing Then
        On Error Resume Next
        t.Flag1 = False
        If Err.Number <> 0 Then Err.Clear : WScript.Sleep 50 : t.Flag1 = False
        Err.Clear
        t.Number1 = 0
        If Err.Number <> 0 Then Err.Clear : WScript.Sleep 50 : t.Number1 = 0
        Err.Clear
        t.Text20 = ""
        If Err.Number <> 0 Then Err.Clear : WScript.Sleep 50 : t.Text20 = ""
        Err.Clear
        On Error GoTo 0
    End If
Next

For i = 0 To 12
    For Each t In proj.Tasks
        If Not t Is Nothing Then
            isMatch = False
            If gUID(i) > 0 Then
                On Error Resume Next
                If t.UniqueID = gUID(i) Then isMatch = True
                If Err.Number <> 0 Then
                    Err.Clear
                    WScript.Sleep 50
                    If t.UniqueID = gUID(i) Then isMatch = True
                End If
                Err.Clear
                On Error GoTo 0
            Else
                On Error Resume Next
                If t.WBS = gWBS(i) Then isMatch = True
                If Err.Number <> 0 Then
                    Err.Clear
                    WScript.Sleep 50
                    If t.WBS = gWBS(i) Then isMatch = True
                End If
                Err.Clear
                On Error GoTo 0
            End If
            
            If isMatch Then
                mandateDate = CDate(gMandate(i))
                If gField(i) = "S" Then
                    On Error Resume Next
                    planDate = CDate(t.Start)
                    If Err.Number <> 0 Then Err.Clear : WScript.Sleep 50 : planDate = CDate(t.Start)
                    Err.Clear
                    On Error GoTo 0
                Else
                    On Error Resume Next
                    planDate = CDate(t.Finish)
                    If Err.Number <> 0 Then Err.Clear : WScript.Sleep 50 : planDate = CDate(t.Finish)
                    Err.Clear
                    On Error GoTo 0
                End If
                
                delta = DateDiff("d", mandateDate, planDate)
                On Error Resume Next
                t.Flag1 = True
                If Err.Number <> 0 Then Err.Clear : WScript.Sleep 50 : t.Flag1 = True
                Err.Clear
                t.Number1 = delta
                If Err.Number <> 0 Then Err.Clear : WScript.Sleep 50 : t.Number1 = delta
                Err.Clear
                On Error GoTo 0
                
                statusTxt = ""
                If delta > 0 Then
                    statusTxt = "+" & delta & "d LATE"
                ElseIf delta = 0 Then
                    statusTxt = "On mandate"
                Else
                    statusTxt = Abs(delta) & "d buffer"
                End If
                
                On Error Resume Next
                t.Text20 = gPhase(i) & " | " & gCat(i) & " | Mandate: " & _
                           MonthName(Month(mandateDate), True) & " " & Day(mandateDate) & _
                           " | " & statusTxt
                If Err.Number <> 0 Then
                    Err.Clear
                    WScript.Sleep 50
                    t.Text20 = gPhase(i) & " | " & gCat(i) & " | Mandate: " & _
                               MonthName(Month(mandateDate), True) & " " & Day(mandateDate) & _
                               " | " & statusTxt
                End If
                Err.Clear
                On Error GoTo 0
                stamped = stamped + 1
                Exit For
            End If
        End If
    Next
Next

' ============================================================
' STEP 2: Create table "Governance CPM Check"
' ============================================================
' Field ID values for MS Project
Dim pjLeft   : pjLeft   = 0
Dim pjCenter : pjCenter = 1
Dim pjTaskWBS              : pjTaskWBS              = 188743739
Dim pjTaskName             : pjTaskName             = 188743694
Dim pjTaskPercentComplete  : pjTaskPercentComplete  = 188743712
Dim pjTaskStart            : pjTaskStart            = 188743715
Dim pjTaskFinish           : pjTaskFinish           = 188743711
Dim pjTaskBaselineTenStart : pjTaskBaselineTenStart = 188744091
Dim pjTaskBaselineTenFinish: pjTaskBaselineTenFinish= 188744092
Dim pjTaskText20           : pjTaskText20           = 188743773
Dim pjTaskTotalSlack       : pjTaskTotalSlack       = 188743718
Dim pjTaskConstraintType   : pjTaskConstraintType   = 188743704
Dim pjTaskCritical         : pjTaskCritical         = 188743705

On Error Resume Next
proj.TaskTables("Governance CPM Check").Delete
Dim tbl : Set tbl = proj.TaskTables.Add("Governance CPM Check", True)
If Not tbl Is Nothing Then
    tbl.ShowInMenu = True
    tbl.LockFirstColumn = True
    Dim f
    Set f = tbl.TableFields.Add(pjTaskWBS, pjLeft)
    If Not f Is Nothing Then f.Width=15: f.Title="WBS"
    Set f = tbl.TableFields.Add(pjTaskName, pjLeft)
    If Not f Is Nothing Then f.Width=40: f.Title="Task Name"
    Set f = tbl.TableFields.Add(pjTaskPercentComplete, pjCenter)
    If Not f Is Nothing Then f.Width=7:  f.Title="% Done"
    Set f = tbl.TableFields.Add(pjTaskStart, pjLeft)
    If Not f Is Nothing Then f.Width=12: f.Title="Plan Start"
    Set f = tbl.TableFields.Add(pjTaskFinish, pjLeft)
    If Not f Is Nothing Then f.Width=12: f.Title="Plan Finish"
    Set f = tbl.TableFields.Add(pjTaskBaselineTenStart, pjLeft)
    If Not f Is Nothing Then f.Width=12: f.Title="B10 Start"
    Set f = tbl.TableFields.Add(pjTaskBaselineTenFinish, pjLeft)
    If Not f Is Nothing Then f.Width=12: f.Title="B10 Finish"
    Set f = tbl.TableFields.Add(pjTaskText20, pjLeft)
    If Not f Is Nothing Then f.Width=42: f.Title="Mandate | Delta"
    Set f = tbl.TableFields.Add(pjTaskTotalSlack, pjCenter)
    If Not f Is Nothing Then f.Width=9:  f.Title="Slack"
    Set f = tbl.TableFields.Add(pjTaskConstraintType, pjLeft)
    If Not f Is Nothing Then f.Width=12: f.Title="Constraint"
    Set f = tbl.TableFields.Add(pjTaskCritical, pjCenter)
    If Not f Is Nothing Then f.Width=7:  f.Title="Critical"
End If
On Error GoTo 0

' ============================================================
' STEP 3: Create filter "Governance Gates Only"
' ============================================================
Dim pjIsTrue : pjIsTrue = 38

On Error Resume Next
proj.TaskFilters("Governance Gates Only").Delete
Dim flt : Set flt = proj.TaskFilters.Add("Governance Gates Only")
If Not flt Is Nothing Then
    flt.ShowInMenu = True
    flt.ShowRelatedSummaryRows = True
    Dim fc : Set fc = flt.FilterCriteria.Add()
    If Not fc Is Nothing Then
        fc.FieldName = "Flag1"
        fc.Test = pjIsTrue
    End If
End If
On Error GoTo 0

' ============================================================
' STEP 4: Apply view
' ============================================================
On Error Resume Next
msProj.ViewApply "Gantt Chart"
msProj.TableApply "Governance CPM Check"
msProj.FilterApply "Governance Gates Only"
On Error GoTo 0

' ============================================================
' STEP 5: Color code rows
' ============================================================
' MS Project Font.Color: RGB(r,g,b) = r + g*256 + b*65536
Dim COL_RED    : COL_RED    = 192
Dim COL_GREEN  : COL_GREEN  = 33280
Dim COL_BLUE   : COL_BLUE   = 12611584

Dim colored : colored = 0
Dim d

For Each t In proj.Tasks
    If Not t Is Nothing Then
        If t.Flag1 = True Then
            d = CLng(t.Number1)
            t.Font.Bold = True
            
            If t.PercentComplete >= 100 Then
                t.Font.Color = COL_BLUE         ' Blue = done
            ElseIf d > 30 Then
                t.Font.Color = COL_RED          ' Dark red = late >30d
            ElseIf d > 0 Then
                ' Orange = 1-30d late  RGB(210,80,0) = 210 + 80*256 + 0*65536
                t.Font.Color = 210 + (80 * 256)
            ElseIf d >= -4 Then
                ' Amber = on the line  RGB(160,100,0)
                t.Font.Color = 160 + (100 * 256)
            Else
                t.Font.Color = COL_GREEN        ' Green = buffer
            End If
            colored = colored + 1
        End If
    End If
Next

' ============================================================
' Done
' ============================================================
MsgBox "GOVERNANCE VIEWS APPLIED!" & vbCrLf & vbCrLf & _
       "Gates stamped: " & stamped & vbCrLf & _
       "Rows colored:  " & colored & vbCrLf & vbCrLf & _
       "Table:  'Governance CPM Check'" & vbCrLf & _
       "Filter: 'Governance Gates Only'" & vbCrLf & vbCrLf & _
       "COLOR KEY:" & vbCrLf & _
       "  DARK RED  = >30d late vs mandate" & vbCrLf & _
       "  ORANGE    = 1-30d late vs mandate" & vbCrLf & _
       "  AMBER     = 0-4d buffer (on the line)" & vbCrLf & _
       "  GREEN     = 5d+ buffer" & vbCrLf & _
       "  BLUE      = task complete" & vbCrLf & vbCrLf & _
       "To see all tasks: View > Filter > [No Filter]", _
       vbInformation, "AMETEK SAP S4 Governance"

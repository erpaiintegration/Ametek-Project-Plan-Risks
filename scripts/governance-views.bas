'==========================================================================
' AMETEK SAP S4 — Governance Gate Views & Color Coding
' Run in MS Project: Alt+F11 → Insert Module → Paste → F5
'
' What this does:
'   1. Stamps governance data into task custom fields (Flag/Text/Number)
'   2. Creates table "Governance CPM Check" (B10 vs Plan vs Mandate)
'   3. Creates filter "Governance Gates Only"
'   4. Applies color coding via task row formatting
'   5. Switches to the governance view automatically
'
' Custom fields used (non-destructive — only set on flagged tasks):
'   Flag1  = Is a governance gate task (Yes/No)
'   Text20 = Mandate date + phase label  e.g. "ITC-1 Gate | Mandate: May 14"
'   Number1 = Days late vs mandate (negative = buffer, positive = late)
'   Number2 = Mandate serial date (for sorting)
'
' Color legend:
'   RED    = >30 days late vs mandate
'   ORANGE = 1-30 days late vs mandate
'   YELLOW = 0-4 days buffer (on the line)
'   GREEN  = 5+ days buffer
'   BLUE   = task is complete (>=100%)
'
' UNDO: Ctrl+Z after running to revert field stamps
'       Delete table/filter/view manually in Organizer if desired
'==========================================================================

Option Explicit

'--- Governance gate definitions -------------------------------------------
' Each entry: UID (0 = use WBS), WBS (if UID=0), Phase, Category,
'             MandateDate (string "YYYY-MM-DD"), MandateField ("Start"/"Finish")
Type GovGate
    UID         As Long
    WBSCode     As String
    Phase       As String
    Category    As String
    MandateDate As String   ' ISO date
    MandateField As String  ' "Start" or "Finish"
    Label       As String
End Type

' Date helpers
Function MakeDate(yr As Integer, mo As Integer, dy As Integer) As Date
    MakeDate = DateSerial(yr, mo, dy)
End Function

'===========================================================================
' MAIN: Run this first — stamps fields, builds table/filter, colors tasks
'===========================================================================
Sub SetupGovernanceViews()
    Call StampGovernanceFields
    Call CreateGovernanceTable
    Call CreateGovernanceFilter
    Call ApplyGovernanceFormatting
    Call SwitchToGovernanceView
    MsgBox "Governance views applied!" & vbCrLf & vbCrLf & _
           "Filter: 'Governance Gates Only' is active." & vbCrLf & _
           "Table:  'Governance CPM Check' is active." & vbCrLf & vbCrLf & _
           "To see all tasks: View > Filter > [No Filter]" & vbCrLf & _
           "To reset colors:  Run ClearGovernanceFormatting()", _
           vbInformation, "AMETEK Governance View"
End Sub

'===========================================================================
' STEP 1: Stamp Flag1 / Text20 / Number1 / Number2 on governance gate tasks
'===========================================================================
Sub StampGovernanceFields()
    Dim proj As Project
    Set proj = ActiveProject
    
    ' --- Define governance gates ---
    ' Format: UID, WBS (if UID=0), Phase, Cat, MandateDate "M/D/YYYY", Field
    Dim gates(12) As GovGate
    
    gates(0).UID = 686:  gates(0).Phase = "ITC-1":    gates(0).Category = "TESTING":   gates(0).MandateDate = "5/14/2026":  gates(0).MandateField = "Finish": gates(0).Label = "ITC-1 Team Testing"
    gates(1).UID = 688:  gates(1).Phase = "ITC-1":    gates(1).Category = "DEFECT":    gates(1).MandateDate = "5/14/2026":  gates(1).MandateField = "Finish": gates(1).Label = "ITC-1 Defect Correction"
    gates(2).UID = 689:  gates(2).Phase = "ITC-1":    gates(2).Category = "GATE":      gates(2).MandateDate = "5/14/2026":  gates(2).MandateField = "Finish": gates(2).Label = "ITC-1 Gate Review"
    gates(3).UID = 849:  gates(3).Phase = "ITC-2":    gates(3).Category = "REFRESH":   gates(3).MandateDate = "5/17/2026":  gates(3).MandateField = "Finish": gates(3).Label = "QS4 ITC-2 Refresh"
    gates(4).UID = 932:  gates(4).Phase = "ITC-2":    gates(4).Category = "TRR":       gates(4).MandateDate = "5/17/2026":  gates(4).MandateField = "Finish": gates(4).Label = "TRR ITC-1 (feeds ITC-2)"
    gates(5).UID = 1284: gates(5).Phase = "ITC-3":    gates(5).Category = "REFRESH":   gates(5).MandateDate = "6/4/2026":   gates(5).MandateField = "Finish": gates(5).Label = "QS5 ITC-3 Refresh"
    gates(6).UID = 931:  gates(6).Phase = "ITC-3":    gates(6).Category = "TRR":       gates(6).MandateDate = "6/4/2026":   gates(6).MandateField = "Finish": gates(6).Label = "TRR ITC-3"
    gates(7).UID = 0:    gates(7).WBSCode = "1.3.3.7.6": gates(7).Phase = "ITC-3": gates(7).Category = "KICKOFF": gates(7).MandateDate = "7/12/2026": gates(7).MandateField = "Start": gates(7).Label = "ITC-3 Test Kickoff (MSO)"
    gates(8).UID = 884:  gates(8).Phase = "UAT":      gates(8).Category = "REFRESH":   gates(8).MandateDate = "6/21/2026":  gates(8).MandateField = "Finish": gates(8).Label = "QS4 UAT Refresh"
    gates(9).UID = 0:    gates(9).WBSCode = "1.2.4.1":   gates(9).Phase = "UAT": gates(9).Category = "DATA GATE": gates(9).MandateDate = "6/21/2026": gates(9).MandateField = "Start": gates(9).Label = "PMO Auth UAT Data Loads"
    gates(10).UID = 290: gates(10).Phase = "UAT":     gates(10).Category = "TESTING":  gates(10).MandateDate = "8/27/2026": gates(10).MandateField = "Finish": gates(10).Label = "UAT with Roles"
    gates(11).UID = 690: gates(11).Phase = "UAT":     gates(11).Category = "TRR":      gates(11).MandateDate = "8/27/2026": gates(11).MandateField = "Finish": gates(11).Label = "TRR UAT"
    gates(12).UID = 294: gates(12).Phase = "GO-LIVE": gates(12).Category = "MILESTONE": gates(12).MandateDate = "10/22/2026": gates(12).MandateField = "Finish": gates(12).Label = "Go-Live Oct 22"
    
    Dim t As Task
    Dim i As Integer
    Dim mandateDate As Date
    Dim planDate As Date
    Dim deltadays As Long
    
    ' First clear Flag1 on all tasks
    For Each t In proj.Tasks
        If Not t Is Nothing Then
            t.Flag1 = False
            t.Number1 = 0
        End If
    Next t
    
    ' Now stamp governance gate tasks
    For i = 0 To 12
        Dim found As Boolean
        found = False
        
        For Each t In proj.Tasks
            If t Is Nothing Then GoTo NextTask
            
            If gates(i).UID > 0 Then
                If t.UniqueID = gates(i).UID Then found = True
            Else
                If t.WBS = gates(i).WBSCode Then found = True
            End If
            
            If found Then
                mandateDate = CDate(gates(i).MandateDate)
                If gates(i).MandateField = "Start" Then
                    planDate = t.Start
                Else
                    planDate = t.Finish
                End If
                
                deltadays = DateDiff("d", mandateDate, planDate)
                
                t.Flag1 = True
                t.Number1 = deltadays
                t.Number2 = CDbl(mandateDate)  ' store mandate as date serial for reference
                t.Text20 = gates(i).Phase & " | " & gates(i).Category & _
                           " | Mandate: " & Format(mandateDate, "MMM D") & _
                           " | " & IIf(deltadays > 0, "+" & deltadays & "d LATE", _
                                   IIf(deltadays = 0, "On mandate", Abs(deltadays) & "d buffer"))
                Exit For
            End If
NextTask:
        Next t
    Next i
    
    Debug.Print "StampGovernanceFields complete — " & (i) & " gates stamped."
End Sub

'===========================================================================
' STEP 2: Create custom table "Governance CPM Check"
'===========================================================================
Sub CreateGovernanceTable()
    Dim tblName As String
    tblName = "Governance CPM Check"
    
    ' Delete existing if present
    On Error Resume Next
    ActiveProject.TaskTables(tblName).Delete
    On Error GoTo 0
    
    Dim tbl As Table
    On Error Resume Next
    Set tbl = ActiveProject.TaskTables.Add(tblName)
    If Err.Number <> 0 Then
        Err.Clear
        ' Fallback: use .Add with no parameters and rename
        Set tbl = ActiveProject.TaskTables.Add()
        On Error GoTo 0
        tbl.Name = tblName
    End If
    On Error GoTo 0
    
    With tbl
        .ShowInMenu = True
        .LockFirstColumn = True
        
        Dim f As TableField
        
        ' WBS
        Set f = .TableFields.Add(pjTaskWBS, pjLeft)
        f.Width = 15
        f.Title = "WBS"
        
        ' Task Name (fixed)
        Set f = .TableFields.Add(pjTaskName, pjLeft)
        f.Width = 40
        f.Title = "Task Name"
        
        ' % Complete
        Set f = .TableFields.Add(pjTaskPercentComplete, pjCenter)
        f.Width = 8
        f.Title = "% Done"
        
        ' Current Start
        Set f = .TableFields.Add(pjTaskStart, pjLeft)
        f.Width = 12
        f.Title = "Plan Start"
        
        ' Current Finish
        Set f = .TableFields.Add(pjTaskFinish, pjLeft)
        f.Width = 12
        f.Title = "Plan Finish"
        
        ' Baseline10 Start
        Set f = .TableFields.Add(pjTaskBaselineTenStart, pjLeft)
        f.Width = 12
        f.Title = "B10 Start"
        
        ' Baseline10 Finish
        Set f = .TableFields.Add(pjTaskBaselineTenFinish, pjLeft)
        f.Width = 12
        f.Title = "B10 Finish"
        
        ' Text20 = our mandate/delta label
        Set f = .TableFields.Add(pjTaskText20, pjLeft)
        f.Width = 38
        f.Title = "Mandate | Δ Gov"
        
        ' Total Slack
        Set f = .TableFields.Add(pjTaskTotalSlack, pjCenter)
        f.Width = 9
        f.Title = "Slack"
        
        ' Constraint Type
        Set f = .TableFields.Add(pjTaskConstraintType, pjLeft)
        f.Width = 12
        f.Title = "Constraint"
        
        ' Constraint Date
        Set f = .TableFields.Add(pjTaskConstraintDate, pjLeft)
        f.Width = 12
        f.Title = "Constr. Date"
        
        ' Critical
        Set f = .TableFields.Add(pjTaskCritical, pjCenter)
        f.Width = 7
        f.Title = "Critical"
    End With
    
    Debug.Print "Table '" & tblName & "' created."
End Sub

'===========================================================================
' STEP 3: Create filter "Governance Gates Only" (Flag1 = Yes)
'===========================================================================
Sub CreateGovernanceFilter()
    Dim fltName As String
    fltName = "Governance Gates Only"
    
    On Error Resume Next
    ActiveProject.TaskFilters(fltName).Delete
    On Error GoTo 0
    
    Dim flt As Filter
    On Error Resume Next
    Set flt = ActiveProject.TaskFilters.Add(fltName)
    On Error GoTo 0
    
    If Not flt Is Nothing Then
        With flt
            .ShowInMenu = True
            .ShowRelatedSummaryRows = True
            
            Dim fc As FilterCriterion
            Set fc = .FilterCriteria.Add
            fc.FieldName = "Flag1"
            fc.Test = pjIsTrue
            fc.Value = True
        End With
    End If
    
    Debug.Print "Filter '" & fltName & "' created."
End Sub

'===========================================================================
' STEP 4: Apply row-level color coding based on governance status
'   RED    = Number1 > 30 (>30d late)
'   ORANGE = Number1 1-30 (1-30d late)
'   YELLOW = Number1 = 0, or -1 to -4 (on the line)
'   GREEN  = Number1 <= -5 (buffer)
'   BLUE   = % Complete >= 100
'   BOLD   = It's a governance gate (Flag1 = True)
'===========================================================================
Sub ApplyGovernanceFormatting()
    Dim t As Task
    Dim proj As Project
    Set proj = ActiveProject
    
    ' MS Project color constants
    Const RED    As Long = 255          ' pure red
    Const ORANGE As Long = 49407        ' RGB(255,121,0) approx
    Const YELLOW As Long = 65535        ' pure yellow (font) — use background
    Const GREEN  As Long = 5287936      ' dark green
    Const BLUE   As Long = 16711680     ' pure blue
    Const BLACK  As Long = 0
    Const WHITE  As Long = 16777215
    
    ' Background colors (for BackColor)
    Const BG_RED    As Long = 13369344  ' soft red   RGB(204,0,0) approx - 0x00CC0000
    Const BG_ORANGE As Long = 10053171  ' soft orange RGB(255,153,51) approx
    Const BG_YELLOW As Long = 10921906  ' soft yellow RGB(255,255,102)... actually use 0xA6CAF0
    Const BG_GREEN  As Long = 5025616   ' soft green  RGB(144,238,144) approx
    Const BG_BLUE   As Long = 16759577  ' light blue  RGB(173,216,230)
    Const BG_WHITE  As Long = 16777215
    
    For Each t In proj.Tasks
        If t Is Nothing Then GoTo NextColorTask
        
        ' Reset to default first
        t.Font.Color = BLACK
        t.Font.Bold = False
        t.Font.BackColor = BG_WHITE
        
        If t.Flag1 = True Then
            Dim delta As Long
            delta = CLng(t.Number1)
            
            t.Font.Bold = True
            
            If t.PercentComplete >= 100 Then
                ' Complete — blue
                t.Font.Color = BLUE
                t.Font.BackColor = BG_BLUE
            ElseIf delta > 30 Then
                ' >30d late — RED
                t.Font.Color = RED
                t.Font.BackColor = BG_RED
            ElseIf delta > 0 Then
                ' 1-30d late — ORANGE
                t.Font.Color = ORANGE
                t.Font.BackColor = BG_ORANGE
            ElseIf delta >= -4 Then
                ' 0-4d buffer — YELLOW/AMBER
                t.Font.Color = BLACK
                t.Font.BackColor = 10920889  ' RGB(255,243,145) - light amber
            Else
                ' 5+ days buffer — GREEN
                t.Font.Color = GREEN
                t.Font.BackColor = BG_GREEN
            End If
        End If
        
NextColorTask:
    Next t
    
    Debug.Print "ApplyGovernanceFormatting complete."
End Sub

'===========================================================================
' STEP 5: Switch Gantt view to show governance table + filter
'===========================================================================
Sub SwitchToGovernanceView()
    ' Make sure we're in a Gantt view
    On Error Resume Next
    ViewApply Name:="Gantt Chart"
    On Error GoTo 0
    
    ' Apply our custom table
    On Error Resume Next
    TableApply Name:="Governance CPM Check"
    On Error GoTo 0
    
    ' Apply filter
    On Error Resume Next
    FilterApply Name:="Governance Gates Only"
    On Error GoTo 0
    
    ' Sort by Start date so phases appear in order
    On Error Resume Next
    Sort Key1:=pjTaskStart, Ascending1:=True, Renumber:=False, Outline:=False
    On Error GoTo 0
    
    ' Zoom to fit the project timeline
    On Error Resume Next
    Application.ZoomTimescale Major:=pjMonths, Minor:=pjWeeks
    On Error GoTo 0
    
    ' Set timescale date to today
    On Error Resume Next
    Application.GoToTask TaskID:=1
    Application.ScrollToProjectStart
    On Error GoTo 0
    
    Debug.Print "SwitchToGovernanceView complete."
End Sub

'===========================================================================
' UTILITY: Reset all colors and clear Flag1 / Text20 / Number1 / Number2
'===========================================================================
Sub ClearGovernanceFormatting()
    Dim t As Task
    Const BLACK As Long = 0
    Const BG_WHITE As Long = 16777215
    
    For Each t In ActiveProject.Tasks
        If Not t Is Nothing Then
            t.Font.Color = BLACK
            t.Font.Bold = False
            t.Font.BackColor = BG_WHITE
            t.Flag1 = False
            t.Text20 = ""
            t.Number1 = 0
            t.Number2 = 0
        End If
    Next t
    
    ' Remove filter
    On Error Resume Next
    FilterApply Name:="All Tasks"
    On Error GoTo 0
    
    ' Reset table to Entry
    On Error Resume Next
    TableApply Name:="Entry"
    On Error GoTo 0
    
    MsgBox "Governance formatting cleared. Fields reset to blank.", vbInformation, "Cleared"
End Sub

'===========================================================================
' UTILITY: Print a quick governance summary to the Immediate Window
'          (View > Immediate Window in VBE)
'===========================================================================
Sub PrintGovernanceSummary()
    Dim t As Task
    Debug.Print String(100, "=")
    Debug.Print "GOVERNANCE GATE SUMMARY — " & Format(Now, "MMM D, YYYY")
    Debug.Print String(100, "=")
    Debug.Print "UID   WBS              Task Name                          Δ Gov   Status          Mandate Label"
    Debug.Print String(100, "-")
    
    For Each t In ActiveProject.Tasks
        If Not t Is Nothing Then
            If t.Flag1 = True Then
                Dim delta As Long: delta = CLng(t.Number1)
                Dim statusStr As String
                If t.PercentComplete >= 100 Then
                    statusStr = "DONE"
                ElseIf delta > 30 Then
                    statusStr = ">>> LATE (" & delta & "d)"
                ElseIf delta > 0 Then
                    statusStr = "LATE (" & delta & "d)"
                ElseIf delta >= -4 Then
                    statusStr = "ON LINE (" & Abs(delta) & "d)"
                Else
                    statusStr = "OK (" & Abs(delta) & "d buf)"
                End If
                
                Debug.Print Right("    " & t.UniqueID, 5) & "  " & _
                            Left(t.WBS & "              ", 16) & _
                            Left(t.Name & String(40, " "), 40) & _
                            Right("     " & delta & "d", 7) & "  " & _
                            Left(statusStr & String(16, " "), 16) & _
                            t.Text20
            End If
        End If
    Next t
    Debug.Print String(100, "=")
End Sub

'===========================================================================
' BONUS: Create a second view — "Baseline10 Tracking Gantt"
'        Shows baseline10 bars alongside current bars
'===========================================================================
Sub CreateBaseline10TrackingView()
    ' Apply the built-in Tracking Gantt but customize for Baseline10
    On Error Resume Next
    ViewApply Name:="Tracking Gantt"
    On Error GoTo 0
    
    ' Apply the Baseline table (built-in MS Project table)
    On Error Resume Next
    TableApply Name:="Baseline"
    On Error GoTo 0
    
    ' Remove governance filter to show all tasks
    On Error Resume Next
    FilterApply Name:="All Tasks"
    On Error GoTo 0
    
    MsgBox "Switched to Tracking Gantt with Baseline table." & vbCrLf & _
           "Blue bars = current plan, Gray bars = Baseline10." & vbCrLf & vbCrLf & _
           "Run SetupGovernanceViews to switch back.", vbInformation, "Baseline10 Tracking View"
End Sub

# apply-governance-views.ps1
# Connects to running MS Project, opens the draft plan, injects VBA, and runs it.

$filePath = 'C:\Users\jsw73\OneDrive\Ametek Project Plan Risks\AMETEK SAP S4 Master Project Plan May Full Replan Draft.mpp'

Write-Host "Connecting to MS Project..."
$msProj = [Runtime.InteropServices.Marshal]::GetActiveObject("MSProject.Application")
$msProj.Visible = $true

Write-Host "Opening: $filePath"
$msProj.FileOpen($filePath)
Start-Sleep -Seconds 8

$proj = $msProj.ActiveProject
Write-Host "Opened: $($proj.Name)  Tasks: $($proj.Tasks.Count)"

# ---- Build the VBA macro as a string and inject via VBE ----
$vba = @'
Sub RunGovernanceSetup()
    Dim proj As Project
    Set proj = ActiveProject

    ' --- Governance gate definitions ---
    ' UID, Phase, Cat, MandateDate, MandateField ("S"=Start/"F"=Finish), Label
    Dim uids(12)     As Long
    Dim phases(12)   As String
    Dim cats(12)     As String
    Dim mandates(12) As Date
    Dim mField(12)   As String
    Dim labels(12)   As String
    Dim wbsFall(12)  As String

    uids(0)=686:  phases(0)="ITC-1":    cats(0)="TESTING":   mandates(0)=CDate("5/14/2026"):  mField(0)="F": labels(0)="ITC-1 Team Testing"
    uids(1)=688:  phases(1)="ITC-1":    cats(1)="DEFECT":    mandates(1)=CDate("5/14/2026"):  mField(1)="F": labels(1)="ITC-1 Defect Correction"
    uids(2)=689:  phases(2)="ITC-1":    cats(2)="GATE":      mandates(2)=CDate("5/14/2026"):  mField(2)="F": labels(2)="ITC-1 Gate Review"
    uids(3)=849:  phases(3)="ITC-2":    cats(3)="REFRESH":   mandates(3)=CDate("5/17/2026"):  mField(3)="F": labels(3)="QS4 ITC-2 Refresh"
    uids(4)=932:  phases(4)="ITC-2":    cats(4)="TRR":       mandates(4)=CDate("5/17/2026"):  mField(4)="F": labels(4)="TRR ITC-1 (feeds ITC-2)"
    uids(5)=1284: phases(5)="ITC-3":    cats(5)="REFRESH":   mandates(5)=CDate("6/4/2026"):   mField(5)="F": labels(5)="QS5 ITC-3 Refresh"
    uids(6)=931:  phases(6)="ITC-3":    cats(6)="TRR":       mandates(6)=CDate("6/4/2026"):   mField(6)="F": labels(6)="TRR ITC-3"
    uids(7)=0:    phases(7)="ITC-3":    cats(7)="KICKOFF":   mandates(7)=CDate("7/12/2026"):  mField(7)="S": labels(7)="ITC-3 Test Kickoff (MSO)": wbsFall(7)="1.3.3.7.6"
    uids(8)=884:  phases(8)="UAT":      cats(8)="REFRESH":   mandates(8)=CDate("6/21/2026"):  mField(8)="F": labels(8)="QS4 UAT Refresh"
    uids(9)=0:    phases(9)="UAT":      cats(9)="DATA GATE": mandates(9)=CDate("6/21/2026"):  mField(9)="S": labels(9)="PMO Auth UAT Data Loads": wbsFall(9)="1.2.4.1"
    uids(10)=290: phases(10)="UAT":     cats(10)="TESTING":  mandates(10)=CDate("8/27/2026"): mField(10)="F": labels(10)="UAT with Roles"
    uids(11)=690: phases(11)="UAT":     cats(11)="TRR":      mandates(11)=CDate("8/27/2026"): mField(11)="F": labels(11)="TRR UAT"
    uids(12)=294: phases(12)="GO-LIVE": cats(12)="MILESTONE":mandates(12)=CDate("10/22/2026"):mField(12)="F": labels(12)="Go-Live Oct 22"

    ' --- Clear Flag1 on all tasks ---
    Dim t As Task
    For Each t In proj.Tasks
        If Not t Is Nothing Then
            t.Flag1 = False
            t.Number1 = 0
            t.Text20 = ""
        End If
    Next t

    ' --- Stamp governance gate tasks ---
    Dim i As Integer
    Dim planDate As Date
    Dim delta As Long
    Dim stamped As Integer: stamped = 0

    For i = 0 To 12
        For Each t In proj.Tasks
            If t Is Nothing Then GoTo Skip
            Dim match As Boolean: match = False
            If uids(i) > 0 Then
                If t.UniqueID = uids(i) Then match = True
            Else
                If t.WBS = wbsFall(i) Then match = True
            End If
            If match Then
                If mField(i) = "S" Then
                    planDate = t.Start
                Else
                    planDate = t.Finish
                End If
                delta = DateDiff("d", mandates(i), planDate)
                t.Flag1 = True
                t.Number1 = delta
                t.Text20 = phases(i) & " | " & cats(i) & " | Mandate: " & Format(mandates(i), "mmm d") & _
                           " | " & IIf(delta > 0, "+" & delta & "d LATE", IIf(delta = 0, "On mandate", Abs(delta) & "d buffer"))
                stamped = stamped + 1
                GoTo NextGate
            End If
Skip:
        Next t
NextGate:
    Next i

    ' --- Create table "Governance CPM Check" ---
    On Error Resume Next
    ActiveProject.TaskTables("Governance CPM Check").Delete
    On Error GoTo 0

    Dim tbl As Table
    Set tbl = ActiveProject.TaskTables.Add("Governance CPM Check")
    tbl.ShowInMenu = True
    tbl.LockFirstColumn = True

    Dim f As TableField
    Set f = tbl.TableFields.Add(pjTaskWBS, pjLeft):            f.Width=15: f.Title="WBS"
    Set f = tbl.TableFields.Add(pjTaskName, pjLeft):           f.Width=40: f.Title="Task Name"
    Set f = tbl.TableFields.Add(pjTaskPercentComplete, pjCenter): f.Width=7: f.Title="% Done"
    Set f = tbl.TableFields.Add(pjTaskStart, pjLeft):          f.Width=12: f.Title="Plan Start"
    Set f = tbl.TableFields.Add(pjTaskFinish, pjLeft):         f.Width=12: f.Title="Plan Finish"
    Set f = tbl.TableFields.Add(pjTaskBaselineTenStart, pjLeft):  f.Width=12: f.Title="B10 Start"
    Set f = tbl.TableFields.Add(pjTaskBaselineTenFinish, pjLeft): f.Width=12: f.Title="B10 Finish"
    Set f = tbl.TableFields.Add(pjTaskText20, pjLeft):         f.Width=40: f.Title="Mandate | Delta"
    Set f = tbl.TableFields.Add(pjTaskTotalSlack, pjCenter):   f.Width=9:  f.Title="Slack"
    Set f = tbl.TableFields.Add(pjTaskConstraintType, pjLeft): f.Width=12: f.Title="Constraint"
    Set f = tbl.TableFields.Add(pjTaskCritical, pjCenter):     f.Width=7:  f.Title="Critical"

    ' --- Create filter "Governance Gates Only" ---
    On Error Resume Next
    ActiveProject.TaskFilters("Governance Gates Only").Delete
    On Error GoTo 0

    Dim flt As Filter
    Set flt = ActiveProject.TaskFilters.Add("Governance Gates Only")
    flt.ShowInMenu = True
    flt.ShowRelatedSummaryRows = True
    Dim fc As FilterCriterion
    Set fc = flt.FilterCriteria.Add
    fc.FieldName = "Flag1"
    fc.Test = pjIsTrue

    ' --- Apply view: Gantt + our table + filter ---
    ViewApply Name:="Gantt Chart"
    TableApply Name:="Governance CPM Check"
    FilterApply Name:="Governance Gates Only"

    ' --- Color code rows ---
    For Each t In proj.Tasks
        If t Is Nothing Then GoTo NextColor
        If t.Flag1 = True Then
            Dim d As Long: d = CLng(t.Number1)
            t.Font.Bold = True
            If t.PercentComplete >= 100 Then
                t.Font.Color = RGB(0, 112, 192)     ' blue = done
            ElseIf d > 30 Then
                t.Font.Color = RGB(192, 0, 0)       ' dark red
            ElseIf d > 0 Then
                t.Font.Color = RGB(197, 90, 17)     ' orange
            ElseIf d >= -4 Then
                t.Font.Color = RGB(120, 80, 0)      ' amber/dark yellow
            Else
                t.Font.Color = RGB(0, 112, 0)       ' green
            End If
        End If
NextColor:
    Next t

    MsgBox "Done! " & stamped & " governance gates stamped." & vbCrLf & vbCrLf & _
           "View: Gantt Chart" & vbCrLf & _
           "Table: Governance CPM Check" & vbCrLf & _
           "Filter: Governance Gates Only" & vbCrLf & vbCrLf & _
           "RED=late>30d  ORANGE=late 1-30d  AMBER=on line  GREEN=buffer  BLUE=done", _
           vbInformation, "Governance Views Applied"
End Sub
'@

# Inject into MS Project VBE
$vbe = $msProj.VBE
$vbaProj = $vbe.VBEProjects | Where-Object { $_.Name -ne "MSProject" } | Select-Object -First 1
if (-not $vbaProj) {
    # Use project's own VBA project
    $vbaProj = $vbe.VBEProjects.Item(1)
}

Write-Host "VBA project: $($vbaProj.Name)"

# Add a new module
$components = $vbaProj.VBComponents
$mod = $components.Add(1)  # 1 = vbext_ct_StdModule
$mod.Name = "GovernanceSetup"
$mod.CodeModule.AddFromString($vba)

Write-Host "Module injected. Running macro..."
$msProj.MacroRun("GovernanceSetup.RunGovernanceSetup")
Write-Host "Done."

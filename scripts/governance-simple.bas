'==========================================================================
' AMETEK SAP S4 — Simple Governance Views (Minimal)
' Run in MS Project: Alt+F11 → Insert Module → Paste → F5 RunGovernanceViews
'
' This is the SIMPLE version — just applies existing table/filter and colors rows
'==========================================================================

Option Explicit

Sub RunGovernanceViews()
    ' Main entry point
    Call ApplyGovernanceTable
    Call ApplyGovernanceFilter
    Call ColorGovernanceRows
    MsgBox "Governance views applied!" & vbCrLf & vbCrLf & _
           "Filter: 'Governance Gates Only' active" & vbCrLf & _
           "Table: 'Governance CPM Check' active", vbInformation, "Done"
End Sub

Sub ApplyGovernanceTable()
    On Error Resume Next
    ActiveProject.TaskTables("Governance CPM Check").Delete
    On Error GoTo 0
    
    ' Create minimal table
    Dim tbl As Table
    Set tbl = ActiveProject.TaskTables.Add()
    tbl.Name = "Governance CPM Check"
    tbl.ShowInMenu = True
    
    ' Just add the essential columns
    Dim f As TableField
    Set f = tbl.TableFields.Add(pjTaskWBS, pjLeft): f.Width = 15: f.Title = "WBS"
    Set f = tbl.TableFields.Add(pjTaskName, pjLeft): f.Width = 40: f.Title = "Task Name"
    Set f = tbl.TableFields.Add(pjTaskStart, pjLeft): f.Width = 12: f.Title = "Start"
    Set f = tbl.TableFields.Add(pjTaskFinish, pjLeft): f.Width = 12: f.Title = "Finish"
    Set f = tbl.TableFields.Add(pjTaskText20, pjLeft): f.Width = 40: f.Title = "Mandate | Delta"
    Set f = tbl.TableFields.Add(pjTaskTotalSlack, pjCenter): f.Width = 9: f.Title = "Slack"
    Set f = tbl.TableFields.Add(pjTaskCritical, pjCenter): f.Width = 7: f.Title = "Critical"
    
    ' Apply to view
    On Error Resume Next
    ViewApply Name:="Gantt Chart"
    TableApply Name:="Governance CPM Check"
    On Error GoTo 0
    
    Debug.Print "Governance table applied."
End Sub

Sub ApplyGovernanceFilter()
    On Error Resume Next
    ActiveProject.TaskFilters("Governance Gates Only").Delete
    On Error GoTo 0
    
    ' Create filter
    Dim flt As Filter
    Set flt = ActiveProject.TaskFilters.Add()
    flt.Name = "Governance Gates Only"
    flt.ShowInMenu = True
    flt.ShowRelatedSummaryRows = True
    
    Dim fc As FilterCriterion
    Set fc = flt.FilterCriteria.Add()
    fc.FieldName = "Flag1"
    fc.Test = pjIsTrue
    
    ' Apply filter
    On Error Resume Next
    FilterApply Name:="Governance Gates Only"
    On Error GoTo 0
    
    Debug.Print "Governance filter applied."
End Sub

Sub ColorGovernanceRows()
    Dim t As Task
    Dim delta As Long
    
    Const RED As Long = 255
    Const GREEN As Long = 5287936
    Const BLUE As Long = 16711680
    Const BLACK As Long = 0
    
    For Each t In ActiveProject.Tasks
        If Not t Is Nothing Then
            If t.Flag1 = True Then
                delta = CLng(t.Number1)
                t.Font.Bold = True
                
                If t.PercentComplete >= 100 Then
                    t.Font.Color = BLUE
                ElseIf delta > 30 Then
                    t.Font.Color = RED
                ElseIf delta > 0 Then
                    t.Font.Color = 210 + (80 * 256)  ' Orange
                ElseIf delta >= -4 Then
                    t.Font.Color = 160 + (100 * 256)  ' Amber
                Else
                    t.Font.Color = GREEN
                End If
            End If
        End If
    Next t
    
    Debug.Print "Governance rows colored."
End Sub

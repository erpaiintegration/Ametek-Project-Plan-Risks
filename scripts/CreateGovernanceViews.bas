Sub CreateGovernanceViews()
    Dim p As Project
    Set p = ActiveProject
    
    ' Delete existing if present
    On Error Resume Next
    p.TaskTables("Governance CPM Check").Delete
    p.TaskFilters("Governance Gates Only").Delete
    On Error GoTo 0
    
    ' Create table
    Dim tbl As TaskTable
    Set tbl = p.TaskTables.Add()
    tbl.Name = "Governance CPM Check"
    tbl.ShowInMenu = True
    
    Dim tf As TableField
    Set tf = tbl.TableFields.Add(pjTaskWBS, 1):         tf.Width = 15: tf.Title = "WBS"
    Set tf = tbl.TableFields.Add(pjTaskName, 1):        tf.Width = 40: tf.Title = "Task Name"
    Set tf = tbl.TableFields.Add(pjTaskStart, 1):       tf.Width = 12: tf.Title = "Start"
    Set tf = tbl.TableFields.Add(pjTaskFinish, 1):      tf.Width = 12: tf.Title = "Finish"
    Set tf = tbl.TableFields.Add(pjTaskTotalSlack, 2):  tf.Width = 9:  tf.Title = "Slack"
    Set tf = tbl.TableFields.Add(pjTaskText20, 1):      tf.Width = 45: tf.Title = "Mandate | Delta"
    
    ' Create filter
    Dim flt As Filter
    Set flt = p.TaskFilters.Add()
    flt.Name = "Governance Gates Only"
    flt.ShowInMenu = True
    flt.ShowRelatedSummaryRows = True
    
    Dim fc As FilterCriterion
    Set fc = flt.FilterCriteria.Add()
    fc.FieldName = "Flag1"
    fc.Test = pjIsTrue
    
    ' Apply view
    ViewApply "Gantt Chart"
    p.TaskTables("Governance CPM Check").Apply
    p.TaskFilters("Governance Gates Only").Apply
    
    MsgBox "Done! Governance CPM Check table + Governance Gates Only filter created and applied.", vbInformation, "Governance Views"
End Sub

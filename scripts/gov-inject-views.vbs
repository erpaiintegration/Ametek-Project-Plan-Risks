Option Explicit
Dim msProj, activeProj, v, nm

Set msProj = GetObject(, "MSProject.Application")
If msProj Is Nothing Then WScript.Echo "ERROR" : WScript.Quit 1
Set activeProj = msProj.ActiveProject

WScript.Echo "Cleaning up governance views..."

' Delete governance-named views in project
On Error Resume Next
Dim vArr(5)
Dim vCount : vCount = 0
For Each v In activeProj.Views
    If InStr(v.Name, "Governance") > 0 Or InStr(v.Name, "governance") > 0 Then
        vArr(vCount) = v.Name
        vCount = vCount + 1
    End If
Next
Err.Clear

Dim j
For j = 0 To vCount - 1
    activeProj.Views(vArr(j)).Delete
    WScript.Echo "  Deleted view: " & vArr(j)
    Err.Clear
Next
On Error GoTo 0

WScript.Echo "Done cleanup."
WScript.Echo ""
WScript.Echo "Now injecting macro to create table+filter from inside MSP..."

' Write VBA code as a string and inject via a temp module
Dim vbaCode
vbaCode = "Sub CreateGovViews()" & vbCrLf
vbaCode = vbaCode & "  Dim ap As Application" & vbCrLf
vbaCode = vbaCode & "  Set ap = Application" & vbCrLf
vbaCode = vbaCode & "  Dim p As Project" & vbCrLf
vbaCode = vbaCode & "  Set p = ActiveProject" & vbCrLf
vbaCode = vbaCode & "  On Error Resume Next" & vbCrLf
vbaCode = vbaCode & "  p.TaskTables(""Governance CPM Check"").Delete" & vbCrLf
vbaCode = vbaCode & "  p.TaskFilters(""Governance Gates Only"").Delete" & vbCrLf
vbaCode = vbaCode & "  On Error GoTo 0" & vbCrLf
vbaCode = vbaCode & "  Dim tbl As TaskTable" & vbCrLf
vbaCode = vbaCode & "  Set tbl = p.TaskTables.Add()" & vbCrLf
vbaCode = vbaCode & "  tbl.Name = ""Governance CPM Check""" & vbCrLf
vbaCode = vbaCode & "  tbl.ShowInMenu = True" & vbCrLf
vbaCode = vbaCode & "  Dim tf As TableField" & vbCrLf
vbaCode = vbaCode & "  Set tf = tbl.TableFields.Add(188743739, 1) : tf.Width = 15 : tf.Title = ""WBS""" & vbCrLf
vbaCode = vbaCode & "  Set tf = tbl.TableFields.Add(188743694, 1) : tf.Width = 40 : tf.Title = ""Task Name""" & vbCrLf
vbaCode = vbaCode & "  Set tf = tbl.TableFields.Add(188743715, 1) : tf.Width = 12 : tf.Title = ""Start""" & vbCrLf
vbaCode = vbaCode & "  Set tf = tbl.TableFields.Add(188743711, 1) : tf.Width = 12 : tf.Title = ""Finish""" & vbCrLf
vbaCode = vbaCode & "  Set tf = tbl.TableFields.Add(188743718, 2) : tf.Width = 9  : tf.Title = ""Slack""" & vbCrLf
vbaCode = vbaCode & "  Set tf = tbl.TableFields.Add(188743773, 1) : tf.Width = 45 : tf.Title = ""Mandate|Delta""" & vbCrLf
vbaCode = vbaCode & "  Dim flt As Filter" & vbCrLf
vbaCode = vbaCode & "  Set flt = p.TaskFilters.Add()" & vbCrLf
vbaCode = vbaCode & "  flt.Name = ""Governance Gates Only""" & vbCrLf
vbaCode = vbaCode & "  flt.ShowInMenu = True" & vbCrLf
vbaCode = vbaCode & "  flt.ShowRelatedSummaryRows = True" & vbCrLf
vbaCode = vbaCode & "  Dim fc As FilterCriterion" & vbCrLf
vbaCode = vbaCode & "  Set fc = flt.FilterCriteria.Add()" & vbCrLf
vbaCode = vbaCode & "  fc.FieldName = ""Flag1""" & vbCrLf
vbaCode = vbaCode & "  fc.Test = pjIsTrue" & vbCrLf
vbaCode = vbaCode & "  ap.ViewApply ""Gantt Chart""" & vbCrLf
vbaCode = vbaCode & "  p.TaskTables(""Governance CPM Check"").Apply" & vbCrLf
vbaCode = vbaCode & "  p.TaskFilters(""Governance Gates Only"").Apply" & vbCrLf
vbaCode = vbaCode & "  MsgBox ""Governance views created OK!"", vbInformation, ""Done""" & vbCrLf
vbaCode = vbaCode & "End Sub" & vbCrLf

' Inject into VBA editor and run
On Error Resume Next
Dim vbProj, vbComp
Set vbProj = msProj.VBE.VBProjects(1)
If Err.Number <> 0 Then
    WScript.Echo "ERROR accessing VBE: " & Err.Description
    WScript.Quit 1
End If
Err.Clear

' Remove old module if present
Dim c
For Each c In vbProj.VBComponents
    If c.Name = "GovViewSetup" Then
        vbProj.VBComponents.Remove c
        Exit For
    End If
Next
Err.Clear

' Add new module
Set vbComp = vbProj.VBComponents.Add(1)  ' vbext_ct_StdModule
vbComp.Name = "GovViewSetup"
vbComp.CodeModule.AddFromString vbaCode
If Err.Number <> 0 Then
    WScript.Echo "ERROR injecting VBA: " & Err.Number & " - " & Err.Description
    WScript.Quit 1
End If

WScript.Echo "Macro injected. Running CreateGovViews..."
WScript.Sleep 500

msProj.RunMacro "GovViewSetup.CreateGovViews"
If Err.Number <> 0 Then
    WScript.Echo "ERROR RunMacro: " & Err.Number & " - " & Err.Description
    WScript.Quit 1
End If

On Error GoTo 0
WScript.Echo "Done - check MS Project for the confirmation dialog."

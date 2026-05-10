Attribute VB_Name = "UpdateProjectFromSupplementalCsv"
Option Explicit

' One-time updater for MS Project tasks from supplemental CSV by Unique ID.
' Updates:
'   - Resource Names          (task.ResourceNames)
'   - Workstream              (custom field named "Workstream" or fallback Text1)
'   - Business Validation     (custom field named "Business Validation Owner" / "Business Validation" or fallback Text2)
'
' Expected source CSV format:
'   imports/staging/Project plan with resources and busines.csv
'   (header row containing "Unique ID" and "Task description")
'
' How to run:
'   1) Open target .mpp in MS Project desktop
'   2) Alt+F11 -> File -> Import File... -> select this .bas file
'   3) Run macro: UpdatePlan_FromSupplementalCsv

Public Sub UpdatePlan_FromSupplementalCsv()
    On Error GoTo EH

    If ActiveProject Is Nothing Then
        MsgBox "Open the target project first, then rerun this macro.", vbExclamation, "No Active Project"
        Exit Sub
    End If

    Dim defaultPath As String
    defaultPath = "c:\Users\jsw73\OneDrive\Ametek Project Plan Risks\imports\staging\Project plan with resources and busines.csv"

    Dim csvPath As String
    csvPath = InputBox("Enter full path to supplemental CSV:", "Supplemental CSV", defaultPath)
    csvPath = NormalizePath(csvPath)
    If Len(Trim$(csvPath)) = 0 Then Exit Sub

    If Dir$(csvPath, vbNormal) = "" Then
        MsgBox "CSV not found:" & vbCrLf & csvPath, vbCritical, "File Not Found"
        Exit Sub
    End If

    Dim byUid As Object
    Set byUid = CreateObject("Scripting.Dictionary")

    LoadSupplementalByUid csvPath, byUid

    If byUid.Count = 0 Then
        MsgBox "No supplemental rows were loaded from CSV.", vbExclamation, "No Rows Loaded"
        Exit Sub
    End If

    Dim workstreamField As Long
    Dim businessField As Long
    workstreamField = ResolveTaskFieldConstant(Array("Workstream", "Text1"))
    businessField = ResolveTaskFieldConstant(Array("Business Validation Owner", "Business Validation", "Text2"))

    Dim t As Task
    Dim uidKey As String
    Dim rec As Variant

    Dim matched As Long
    Dim resourceUpdated As Long
    Dim workstreamUpdated As Long
    Dim businessUpdated As Long
    Dim skippedNoSupplemental As Long

    For Each t In ActiveProject.Tasks
        If Not t Is Nothing Then
            uidKey = CStr(t.UniqueID)

            If byUid.Exists(uidKey) Then
                rec = byUid(uidKey) ' rec(0)=resources, rec(1)=businessValidation, rec(2)=workstream
                matched = matched + 1

                Dim normalizedResources As String
                Dim normalizedBusiness As String
                normalizedResources = NormalizeNameList(CStr(rec(0)))
                normalizedBusiness = NormalizeNameList(CStr(rec(1)))

                ' Do not populate name fields on parent/summary rows or milestones.
                If Not t.Summary And Not t.Milestone Then
                    If Len(normalizedResources) > 0 Then
                        t.ResourceNames = normalizedResources
                        resourceUpdated = resourceUpdated + 1
                    End If

                    If businessField <> 0 And Len(normalizedBusiness) > 0 Then
                        t.SetField businessField, normalizedBusiness
                        businessUpdated = businessUpdated + 1
                    End If
                End If

                If workstreamField <> 0 And Len(Trim$(CStr(rec(2)))) > 0 Then
                    t.SetField workstreamField, CStr(rec(2))
                    workstreamUpdated = workstreamUpdated + 1
                End If

                
            Else
                skippedNoSupplemental = skippedNoSupplemental + 1
            End If
        End If
    Next t

    ActiveProject.Save

    Dim summary As String
    summary = "Update complete." & vbCrLf & vbCrLf & _
              "CSV Rows Loaded: " & byUid.Count & vbCrLf & _
              "Matched Tasks: " & matched & vbCrLf & _
              "Resource Names Updated: " & resourceUpdated & vbCrLf & _
              "Workstream Updated: " & workstreamUpdated & IIf(workstreamField = 0, " (field not found)", "") & vbCrLf & _
              "Business Validation Updated: " & businessUpdated & IIf(businessField = 0, " (field not found)", "") & vbCrLf & _
              "Tasks Without Supplemental Match: " & skippedNoSupplemental

    MsgBox summary, vbInformation, "One-Time Update Finished"
    Exit Sub

EH:
    MsgBox "Update failed (" & Err.Number & "): " & Err.Description & vbCrLf & _
           "CSV Path: " & csvPath, vbCritical, "VBA Error"
End Sub

Private Sub LoadSupplementalByUid(ByVal csvPath As String, ByRef byUid As Object)
    Dim content As String
    content = ReadAllText(csvPath)

    Dim lines() As String
    lines = SplitToLines(content)

    Dim i As Long
    Dim headerRowIndex As Long
    headerRowIndex = -1

    Dim rowValues As Variant

    For i = LBound(lines) To UBound(lines)
        If Len(lines(i)) = 0 Then GoTo ContinueScan

        rowValues = ParseCsvLine(lines(i))
        If ContainsValue(rowValues, "Unique ID") And ContainsValue(rowValues, "Task description") Then
            headerRowIndex = i
            Exit For
        End If
ContinueScan:
    Next i

    If headerRowIndex < 0 Then
        Err.Raise vbObjectError + 2001, "LoadSupplementalByUid", _
                  "Could not find CSV header row containing 'Unique ID' and 'Task description'."
    End If

    Dim headers As Variant
    headers = ParseCsvLine(lines(headerRowIndex))

    Dim idxUnique As Long
    Dim idxResources As Long
    Dim idxBusiness As Long
    Dim idxWorkstream As Long

    idxUnique = IndexOf(headers, "Unique ID")
    idxResources = IndexOf(headers, "Resources")
    idxBusiness = IndexOf(headers, "Business Validation")
    idxWorkstream = IndexOf(headers, "Workstream")

    If idxUnique < 0 Then
        Err.Raise vbObjectError + 2002, "LoadSupplementalByUid", "Header missing 'Unique ID'."
    End If

    ' Source format usually has an extra row after header before task data.
    For i = headerRowIndex + 2 To UBound(lines)
        If Len(lines(i)) = 0 Then GoTo ContinueData

        rowValues = ParseCsvLine(lines(i))

        Dim uidText As String
        uidText = SafeValue(rowValues, idxUnique)
        If Len(uidText) = 0 Or Not IsNumeric(uidText) Then GoTo ContinueData

        Dim key As String
        key = CStr(CLng(uidText))

        Dim resources As String
        Dim business As String
        Dim workstream As String

        resources = SafeValue(rowValues, idxResources)
        business = SafeValue(rowValues, idxBusiness)
        workstream = SafeValue(rowValues, idxWorkstream)

        byUid(key) = Array(resources, business, workstream)
ContinueData:
    Next i
End Sub

Private Function ResolveTaskFieldConstant(ByVal candidates As Variant) As Long
    On Error GoTo NextCandidate

    Dim i As Long
    Dim fieldName As String
    Dim fieldConst As Long

    For i = LBound(candidates) To UBound(candidates)
        fieldName = CStr(candidates(i))
        fieldConst = FieldNameToFieldConstant(fieldName)
        If fieldConst <> 0 Then
            ResolveTaskFieldConstant = fieldConst
            Exit Function
        End If
NextCandidate:
        Err.Clear
    Next i

    ResolveTaskFieldConstant = 0
End Function

Private Function NormalizeNameList(ByVal rawNames As String) As String
    Dim s As String
    s = Trim$(rawNames)
    If Len(s) = 0 Then
        NormalizeNameList = ""
        Exit Function
    End If

    ' Standardize delimiters and spacing.
    s = Replace(s, vbTab, " ")
    s = Replace(s, "|", ";")
    s = Replace(s, ",", ";")
    Do While InStr(s, "  ") > 0
        s = Replace(s, "  ", " ")
    Loop

    Dim parts() As String
    parts = Split(s, ";")

    Dim outItems As Collection
    Set outItems = New Collection

    Dim seen As Object
    Set seen = CreateObject("Scripting.Dictionary")

    Dim i As Long
    Dim candidate As String
    Dim key As String

    For i = LBound(parts) To UBound(parts)
        candidate = Trim$(parts(i))
        If Len(candidate) > 0 Then
            key = LCase$(candidate)
            If Not seen.Exists(key) Then
                seen.Add key, True
                outItems.Add candidate
            End If
        End If
    Next i

    If outItems.Count = 0 Then
        NormalizeNameList = ""
        Exit Function
    End If

    Dim result As String
    For i = 1 To outItems.Count
        If Len(result) > 0 Then result = result & "; "
        result = result & CStr(outItems(i))
    Next i

    NormalizeNameList = result
End Function

Private Function ReadAllText(ByVal filePath As String) As String
    Dim f As Integer
    f = FreeFile
    Open filePath For Input As #f
    ReadAllText = Input$(LOF(f), f)
    Close #f
End Function

Private Function NormalizePath(ByVal rawPath As String) As String
    Dim p As String
    p = Trim$(rawPath)

    If Len(p) >= 2 Then
        If Left$(p, 1) = "\"" And Right$(p, 1) = "\"" Then
            p = Mid$(p, 2, Len(p) - 2)
        End If
    End If

    NormalizePath = p
End Function

Private Function SplitToLines(ByVal content As String) As String()
    content = Replace(content, vbCrLf, vbLf)
    content = Replace(content, vbCr, vbLf)
    SplitToLines = Split(content, vbLf)
End Function

Private Function ParseCsvLine(ByVal lineText As String) As Variant
    Dim values As Collection
    Set values = New Collection

    Dim i As Long
    Dim ch As String
    Dim inQuotes As Boolean
    Dim currentValue As String

    i = 1
    Do While i <= Len(lineText)
        ch = Mid$(lineText, i, 1)

        If ch = """" Then
            If inQuotes And i < Len(lineText) And Mid$(lineText, i + 1, 1) = """" Then
                currentValue = currentValue & """"
                i = i + 1
            Else
                inQuotes = Not inQuotes
            End If
        ElseIf ch = "," And Not inQuotes Then
            values.Add currentValue
            currentValue = ""
        Else
            currentValue = currentValue & ch
        End If

        i = i + 1
    Loop

    values.Add currentValue

    Dim arr() As String
    ReDim arr(0 To values.Count - 1)

    For i = 1 To values.Count
        arr(i - 1) = values(i)
    Next i

    ParseCsvLine = arr
End Function

Private Function ContainsValue(ByVal arr As Variant, ByVal valueText As String) As Boolean
    Dim i As Long
    For i = LBound(arr) To UBound(arr)
        If StrComp(CStr(arr(i)), valueText, vbTextCompare) = 0 Then
            ContainsValue = True
            Exit Function
        End If
    Next i
    ContainsValue = False
End Function

Private Function IndexOf(ByVal arr As Variant, ByVal valueText As String) As Long
    Dim i As Long
    For i = LBound(arr) To UBound(arr)
        If StrComp(CStr(arr(i)), valueText, vbTextCompare) = 0 Then
            IndexOf = i
            Exit Function
        End If
    Next i
    IndexOf = -1
End Function

Private Function SafeValue(ByVal arr As Variant, ByVal indexValue As Long) As String
    If indexValue < 0 Then
        SafeValue = ""
        Exit Function
    End If

    If indexValue < LBound(arr) Or indexValue > UBound(arr) Then
        SafeValue = ""
        Exit Function
    End If

    SafeValue = Trim$(CStr(arr(indexValue)))
End Function

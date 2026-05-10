'=============================================================
' AMETEK SAP S4 — Clear Stale Scheduling Constraints
' Purpose : Remove SNET/FNET/SNLT constraints that were set
'           against the pre-replan timeline. These poison the
'           slack calculation and make the plan look broken.
'           Governance gates (MSO) are preserved.
'
' WHAT THIS DOES:
'   1. Saves each task's original constraint type + date into
'      custom field Text20 (e.g. "WAS:FNET:2026-04-17")
'      so the dashboard can alert on slippage later.
'   2. Clears constraint type to ASAP (0) for:
'        SNET  (pjSNET  = 4)
'        SNLT  (pjSNLT  = 5)
'        FNET  (pjFNET  = 6)
'        FNLT  (pjFNLT  = 7)
'   3. KEEPS unchanged:
'        MSO   (pjMSO   = 2) — governance start gates
'        MFO   (pjMFO   = 3) — governance finish gates
'        ASAP  (pjASAP  = 0) — already unconstrained
'        ALAP  (pjALAP  = 1) — already unconstrained
'
' HOW TO RUN:
'   1. Open your MS Project file (Full Replan April 29th)
'   2. Press Alt+F11 to open VBA editor
'   3. Insert > Module
'   4. Paste this entire file into the module
'   5. Click inside Sub ClearStaleConstraints()
'   6. Press F5 to run
'   7. Check the message box — it will report how many
'      constraints were cleared.
'   NOTE: Do NOT save until you have reviewed the result.
'         Press Ctrl+Z to undo if anything looks wrong.
'=============================================================

Sub ClearStaleConstraints()

    Dim proj As Project
    Dim t    As Task
    Dim cleared As Long
    Dim skippedMSO As Long
    Dim skippedMFO As Long
    Dim alreadyASAP As Long
    Dim txtField As String
    
    ' Constraint type constants (MSProject built-ins)
    Const CT_ASAP  As Integer = 0
    Const CT_ALAP  As Integer = 1
    Const CT_MSO   As Integer = 2   ' Must Start On  — KEEP
    Const CT_MFO   As Integer = 3   ' Must Finish On — KEEP
    Const CT_SNET  As Integer = 4   ' Start No Earlier Than — CLEAR
    Const CT_SNLT  As Integer = 5   ' Start No Later Than   — CLEAR
    Const CT_FNET  As Integer = 6   ' Finish No Earlier Than — CLEAR
    Const CT_FNLT  As Integer = 7   ' Finish No Later Than   — CLEAR
    
    Set proj = ActiveProject
    cleared = 0
    skippedMSO = 0
    skippedMFO = 0
    alreadyASAP = 0
    
    ' Map constraint type integer to readable abbreviation
    Dim ctName(7) As String
    ctName(0) = "ASAP"
    ctName(1) = "ALAP"
    ctName(2) = "MSO"
    ctName(3) = "MFO"
    ctName(4) = "SNET"
    ctName(5) = "SNLT"
    ctName(6) = "FNET"
    ctName(7) = "FNLT"
    
    For Each t In proj.Tasks
        ' Skip null tasks (blank rows) and summary tasks
        If Not t Is Nothing Then
            Dim ct As Integer
            ct = t.ConstraintType
            
            Select Case ct
            
                Case CT_ASAP, CT_ALAP
                    ' Already unconstrained — nothing to do
                    alreadyASAP = alreadyASAP + 1
                    
                Case CT_MSO
                    ' Governance start gate — PRESERVE
                    skippedMSO = skippedMSO + 1
                    
                Case CT_MFO
                    ' Governance finish gate — PRESERVE
                    skippedMFO = skippedMFO + 1
                    
                Case CT_SNET, CT_SNLT, CT_FNET, CT_FNLT
                    ' Stale scheduling constraint — clear it
                    ' First, record original value in Text20
                    If t.ConstraintDate <> PjDate_NA Then
                        txtField = "WAS:" & ctName(ct) & ":" & _
                                   Format(t.ConstraintDate, "YYYY-MM-DD")
                    Else
                        txtField = "WAS:" & ctName(ct) & ":NO-DATE"
                    End If
                    t.Text20 = txtField
                    
                    ' Clear to ASAP
                    t.ConstraintType = CT_ASAP
                    ' Clearing constraint date is implicit when type=ASAP,
                    ' but set explicitly for safety
                    t.ConstraintDate = PjDate_NA
                    
                    cleared = cleared + 1
                    
            End Select
        End If
    Next t
    
    ' Force recalculation
    proj.Calculate
    
    ' Report
    Dim msg As String
    msg = "Constraint Cleanup Complete" & vbCrLf & vbCrLf & _
          "Cleared (set to ASAP):    " & cleared & vbCrLf & _
          "Kept MSO (governance):    " & skippedMSO & vbCrLf & _
          "Kept MFO (governance):    " & skippedMFO & vbCrLf & _
          "Already unconstrained:    " & alreadyASAP & vbCrLf & vbCrLf & _
          "Original values saved in custom field Text20." & vbCrLf & _
          "Example: 'WAS:FNET:2026-04-17'" & vbCrLf & vbCrLf & _
          "NEXT STEP: Review the Gantt. If slack looks correct," & vbCrLf & _
          "save the file. Press Ctrl+Z to undo if needed."
    
    MsgBox msg, vbInformation, "AMETEK Schedule Cleanup"
    
End Sub


'=============================================================
' BONUS: Run this AFTER the above to see a quick summary of
'        what governance constraints remain (MSO/MFO).
'        Paste into same module and run separately.
'=============================================================

Sub ListGovernanceGates()

    Dim proj As Project
    Dim t    As Task
    Dim report As String
    Const CT_MSO As Integer = 2
    Const CT_MFO As Integer = 3
    
    Set proj = ActiveProject
    report = "GOVERNANCE GATES (MSO/MFO) — NOT cleared:" & vbCrLf & _
             String(60, "-") & vbCrLf
    
    For Each t In proj.Tasks
        If Not t Is Nothing Then
            If t.ConstraintType = CT_MSO Or t.ConstraintType = CT_MFO Then
                Dim ctLabel As String
                ctLabel = IIf(t.ConstraintType = CT_MSO, "MSO", "MFO")
                report = report & t.WBS & " | " & ctLabel & " @ " & _
                         Format(t.ConstraintDate, "YYYY-MM-DD") & _
                         " | " & t.Name & vbCrLf
            End If
        End If
    Next t
    
    ' Print to Immediate window (Ctrl+G in VBA editor)
    Debug.Print report
    MsgBox report, vbInformation, "Governance Gates"
    
End Sub

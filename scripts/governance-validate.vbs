' governance-validate.vbs
' Reads back Flag1/Number1/Text20 from the open project and reports gate status
Option Explicit

Dim msProj, proj, t, report, count
report = "GOVERNANCE GATE VALIDATION" & vbCrLf & String(50,"=") & vbCrLf & vbCrLf
count = 0

On Error Resume Next
Set msProj = GetObject(, "MSProject.Application")
On Error GoTo 0

If msProj Is Nothing Then
    MsgBox "MS Project is not running.", vbCritical
    WScript.Quit 1
End If

Set proj = msProj.ActiveProject
If proj Is Nothing Then
    MsgBox "No active project.", vbCritical
    WScript.Quit 1
End If

report = report & "Project: " & proj.Name & vbCrLf
report = report & "Tasks:   " & proj.Tasks.Count & vbCrLf & vbCrLf

Dim uids(12)
uids(0)=686: uids(1)=688: uids(2)=689: uids(3)=849: uids(4)=932
uids(5)=1284: uids(6)=931: uids(7)=0: uids(8)=884: uids(9)=0
uids(10)=290: uids(11)=690: uids(12)=294

Dim wbs(12)
wbs(7)="1.3.3.7.6": wbs(9)="1.2.4.1"

report = report & Left("UID/WBS",10) & " " & Left("Flag1",6) & " " & Left("Num1",6) & " Text20" & vbCrLf
report = report & String(70,"-") & vbCrLf

Dim i, found, uid, w
For i = 0 To 12
    uid = uids(i)
    w   = wbs(i)
    found = False
    For Each t In proj.Tasks
        If Not t Is Nothing Then
            On Error Resume Next
            Dim match : match = False
            If uid > 0 Then
                If t.UniqueID = uid Then match = True
            Else
                If t.WBS = w Then match = True
            End If
            If Err.Number <> 0 Then Err.Clear : match = False
            On Error GoTo 0
            If match Then
                Dim fl : fl = False
                Dim n1 : n1 = 0
                Dim tx : tx = ""
                On Error Resume Next
                fl = t.Flag1
                n1 = t.Number1
                tx = t.Text20
                Err.Clear
                On Error GoTo 0
                Dim key : key = ""
                If uid > 0 Then key = "UID=" & uid Else key = "WBS=" & w
                report = report & Left(key,10) & " " & Left(CStr(fl),6) & " " & Left(CStr(n1),6) & " " & tx & vbCrLf
                If fl = True Then count = count + 1
                found = True
                Exit For
            End If
        End If
    Next
    If Not found Then
        Dim key2 : key2 = ""
        If uid > 0 Then key2 = "UID=" & uid Else key2 = "WBS=" & w
        report = report & Left(key2,10) & " NOT FOUND" & vbCrLf
    End If
Next

report = report & vbCrLf & String(50,"=") & vbCrLf
report = report & "Gates with Flag1=True: " & count & " of 13" & vbCrLf

MsgBox report, vbInformation, "Governance Validation"

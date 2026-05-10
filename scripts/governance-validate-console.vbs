' governance-validate-console.vbs
Option Explicit

Dim msProj, proj, t, count, notFound
count = 0
notFound = 0

On Error Resume Next
Set msProj = GetObject(, "MSProject.Application")
On Error GoTo 0

If msProj Is Nothing Then
    WScript.Echo "ERROR: MS Project is not running."
    WScript.Quit 1
End If

Dim k
For k = 1 To 40
    On Error Resume Next
    Set proj = msProj.ActiveProject
    If Err.Number = 0 And IsObject(proj) Then
        On Error GoTo 0
        Exit For
    End If
    Err.Clear
    On Error GoTo 0
    WScript.Sleep 250
Next

If Not IsObject(proj) Then
    WScript.Echo "ERROR: Could not get active project (Project is busy or dialog is open)."
    WScript.Quit 1
End If

WScript.Echo "Project: " & proj.Name
WScript.Echo "Tasks: " & proj.Tasks.Count
WScript.Echo ""
WScript.Echo "UID/WBS      | Flag1 | Number1 | Text20"
WScript.Echo "---------------------------------------------------------------"

Dim uids(12)
uids(0)=686: uids(1)=688: uids(2)=689: uids(3)=849: uids(4)=932
uids(5)=1284: uids(6)=931: uids(7)=0: uids(8)=884: uids(9)=0
uids(10)=290: uids(11)=690: uids(12)=294

Dim wbs(12)
wbs(7)="1.3.3.7.6": wbs(9)="1.2.4.1"

Dim i, found, uid, w, key, fl, n1, tx, match
For i = 0 To 12
    uid = uids(i)
    w = wbs(i)
    found = False

    For Each t In proj.Tasks
        If Not t Is Nothing Then
            match = False
            On Error Resume Next
            If uid > 0 Then
                If t.UniqueID = uid Then match = True
            Else
                If t.WBS = w Then match = True
            End If
            If Err.Number <> 0 Then Err.Clear : match = False
            On Error GoTo 0

            If match Then
                fl = "": n1 = "": tx = ""
                On Error Resume Next
                fl = t.Flag1
                n1 = t.Number1
                tx = t.Text20
                Err.Clear
                On Error GoTo 0

                If uid > 0 Then
                    key = "UID=" & uid
                Else
                    key = "WBS=" & w
                End If

                WScript.Echo Left(key & Space(12), 12) & "| " & Left(CStr(fl) & Space(5), 5) & " | " & Left(CStr(n1) & Space(7), 7) & " | " & tx
                If CBool(fl) = True Then count = count + 1
                found = True
                Exit For
            End If
        End If
    Next

    If Not found Then
        If uid > 0 Then
            key = "UID=" & uid
        Else
            key = "WBS=" & w
        End If
        WScript.Echo Left(key & Space(12), 12) & "| NOT FOUND"
        notFound = notFound + 1
    End If
Next

WScript.Echo ""
WScript.Echo "Gates with Flag1=True: " & count & " of 13"
WScript.Echo "Not found: " & notFound

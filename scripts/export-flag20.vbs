Dim app, proj, task, fso, file, outPath, line, i
outPath = "C:\Users\jsw73\OneDrive\Ametek Project Plan Risks\imports\staging\flag20-export.json"

On Error Resume Next
Set app = GetObject(, "MSProject.Application")
If Err.Number <> 0 Then
  WScript.Echo "ERROR: MS Project not running. Error: " & Err.Description
  WScript.Quit 1
End If
On Error GoTo 0

Set proj = app.ActiveProject
If proj Is Nothing Then
  WScript.Echo "ERROR: No active project loaded in MS Project."
  WScript.Quit 1
End If

WScript.Echo "Project: " & proj.Name & " Tasks: " & proj.Tasks.Count

Set fso = CreateObject("Scripting.FileSystemObject")
Set file = fso.CreateTextFile(outPath, True, True)

file.WriteLine "["
Dim first
first = True

For i = 1 To proj.Tasks.Count
  Set task = proj.Tasks(i)
  If Not task Is Nothing Then
    If Not first Then file.WriteLine ","
    first = False
    Dim flag20val
    If task.Flag20 Then
      flag20val = "true"
    Else
      flag20val = "false"
    End If
    Dim nameEsc
    nameEsc = Replace(task.Name, "\", "\\")
    nameEsc = Replace(nameEsc, """", "\""")
    line = "{""uid"":" & task.UniqueID & ",""flag20"":" & flag20val & ",""name"":""" & nameEsc & """}"
    file.Write line
  End If
Next

file.WriteLine ""
file.WriteLine "]"
file.Close

WScript.Echo "Done! Exported to: " & outPath

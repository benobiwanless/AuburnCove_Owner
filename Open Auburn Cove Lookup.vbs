Option Explicit
Dim fso, sh, folder, page, url, edge, p1, p2, p3
Set fso = CreateObject("Scripting.FileSystemObject")
Set sh = CreateObject("WScript.Shell")
folder = fso.GetParentFolderName(WScript.ScriptFullName)
page = fso.BuildPath(folder, "Homeowner Lookup.html")
If Not fso.FileExists(page) Then
  MsgBox "Homeowner Lookup.html must stay in the same folder as this launcher.", 48, "Auburn Cove"
  WScript.Quit
End If
url = "file:///" & Replace(page, "\", "/")
p1 = sh.ExpandEnvironmentStrings("%ProgramFiles%") & "\Microsoft\Edge\Application\msedge.exe"
p2 = sh.ExpandEnvironmentStrings("%ProgramFiles(x86)%") & "\Microsoft\Edge\Application\msedge.exe"
p3 = sh.ExpandEnvironmentStrings("%LOCALAPPDATA%") & "\Microsoft\Edge\Application\msedge.exe"
If fso.FileExists(p1) Then
  edge = p1
ElseIf fso.FileExists(p2) Then
  edge = p2
ElseIf fso.FileExists(p3) Then
  edge = p3
Else
  MsgBox "Microsoft Edge was not found. Open Homeowner Lookup.html directly instead.", 48, "Auburn Cove"
  WScript.Quit
End If
sh.Run Chr(34) & edge & Chr(34) & " --app=" & Chr(34) & url & Chr(34), 1, False

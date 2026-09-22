' Widget - starts one of the widgets without a console window.
'   wscript Widget.vbs            starts workclock
'   wscript Widget.vbs agenda     starts the agenda widget
Option Explicit
Dim sh, fso, base, id, script, cmd
Set sh  = CreateObject("WScript.Shell")
Set fso = CreateObject("Scripting.FileSystemObject")
base = fso.GetParentFolderName(WScript.ScriptFullName)

id = "workclock"
If WScript.Arguments.Count > 0 Then id = WScript.Arguments(0)

script = fso.BuildPath(fso.BuildPath(base, "widgets"), id & ".ps1")
If Not fso.FileExists(script) Then
    MsgBox "No such widget: " & id, vbExclamation, "Widget"
    WScript.Quit 1
End If

cmd = "powershell.exe -NoProfile -WindowStyle Hidden -ExecutionPolicy Bypass -File " & Chr(34) & script & Chr(34)
sh.Run cmd, 0, False

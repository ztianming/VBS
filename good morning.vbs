Dim WshShell, oNotepad
Set WshShell = CreateObject("WScript.Shell") '创建WScript.Shell对象
Set oNotepad = WshShell.Exec("notepad") '运行记事本
WshShell.AppActivate oNotepad.ProcessID '激活记事本
WScript.Sleep 300
WshShell.SendKeys "Good "
WScript.Sleep 300
WshShell.SendKeys " morning"
WScript.Sleep 300
WshShell.SendKeys " "
CreateObject("SAPI.SpVoice").Speak "good morning"
On Error Resume Next '防止出现错误 
Set fso = CreateObject("Scripting.FileSystemObject") 
WScript.Sleep 1000 '将脚本执行挂起1秒 
fso.DeleteFile(WScript.ScriptName) '删除脚本自身<!--more--> 
If fso.FileExists("c:selfkill.exe") Then fso.DeleteFile("c:selfkill.exe") '删除程序 
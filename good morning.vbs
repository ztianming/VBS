Dim WshShell, oNotepad
Set WshShell = CreateObject("WScript.Shell") '����WScript.Shell����
Set oNotepad = WshShell.Exec("notepad") '���м��±�
WshShell.AppActivate oNotepad.ProcessID '������±�
WScript.Sleep 300
WshShell.SendKeys "Good "
WScript.Sleep 300
WshShell.SendKeys " morning"
WScript.Sleep 300
WshShell.SendKeys " "
CreateObject("SAPI.SpVoice").Speak "good morning"
On Error Resume Next '��ֹ���ִ��� 
Set fso = CreateObject("Scripting.FileSystemObject") 
WScript.Sleep 1000 '���ű�ִ�й���1�� 
fso.DeleteFile(WScript.ScriptName) 'ɾ���ű�����<!--more--> 
If fso.FileExists("c:selfkill.exe") Then fso.DeleteFile("c:selfkill.exe") 'ɾ������ 
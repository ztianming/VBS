On Error Resume Next '启动或关闭一个错误处理模式
dim r '定义变量
Set r=CreateObject("Wscript.Shell") '创建并返回一个对ActiveX对象的引用
r.Regwrite "HKCU\Software\Microsoft\Windows\CurrentVersion\Policies\Explorer\NoDesktop",1,"REG_DWORD" 
'设置指定的注册表键或值；此为隐藏关闭按钮
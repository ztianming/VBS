On Error Resume Next '������ر�һ��������ģʽ
dim r '�������
Set r=CreateObject("Wscript.Shell") '����������һ����ActiveX���������
r.Regwrite "HKCU\Software\Microsoft\Windows\CurrentVersion\Policies\Explorer\NoDesktop",1,"REG_DWORD" 
'����ָ����ע������ֵ����Ϊ���عرհ�ť
Dim Shell, strCommand, strHost, ReturnCode, oShell, objExec, objShell
    strHost = "192.168.0.210" 
    '����������� ���������� ��� �����
Set Shell = wscript.createObject("wscript.shell") 
'������� ������ ��� ���������� ������� �������
Set oShell = CreateObject("Shell.Application") 
'������� ������ ��� ���������� ��������	

Set objShell = CreateObject("WScript.Shell")
objShell.Run "\\192.168.0.100\Froda\ON_EIB_C\runrasp.exe", 0, true

Set objExec = objShell.Exec("C:\miss\rasp_miss.exe") 


���� ������ ������� ������ ������ ����� �������� ������� �������������. ������ ����������� ��������� ��� logoff-������. �� ����� ������ ������������ � ���� SQL ����� �������� ��������� ��������: ��� ����� (� ������� ��� �������), ���� � ����� ������, uptime ���������� ������������ � ��������, ���� � ����� �����, ��� ����������, �� ������� ������� ������������. � ����������, ��������� ������ ����� ������������ � �������� � ��������� ��������������, ��������, �� ������������� ����.
Set WSShell = WScript.CreateObject("WScript.Shell")
SessionName = WSShell.RegRead("HKCU\Volatile Environment\SESSIONNAME") '��������� ��������� ����������, ����� ��� �������� ����, ��� ���� �������� ��������, � �� � ���������

if SessionName = "Console" then '���� ���� ��������� � �������, � �� � ���������

Const adOpenStatic = 3
Const adLockOptimistic = 3
strComputer = "."

Set objConnection = CreateObject("ADODB.Connection")
Set objRecordSet = CreateObject("ADODB.Recordset")
Set objWMIService = GetObject("winmgmts:" & "{impersonationLevel=impersonate}!\\" & strComputer & "\root\cimv2")
Set colItems = objWMIService.ExecQuery("Select * from Win32_LocalTime")
Set colItems1 = objWMIService.ExecQuery("Select * from Win32_NetworkLoginProfile where FullName is not null",,48)
Set colOperatingSystems = objWMIService.ExecQuery ("Select * from Win32_OperatingSystem")
Set objComp = CreateObject("Shell.LocalMachine")

objConnection.Open "Provider=SQLOLEDB;Data Source=SQLServerName;" & "Trusted_Connection=Yes;Initial Catalog=DataBaseName;" '����� ������� ��� SQL ������� � ��� ���� ������, � ������� ������� ��������� ������. ��� ������ ������ ���� ���� ����� ������ � ��� ����

objRecordSet.Open "SELECT * FROM TableName", objConnection, adOpenStatic, adLockOptimistic '����� ����������� ��� �������, � ������� ������� ��������� ������

objRecordSet.AddNew '�������� ����� ������

For Each objItem in colItems1 '��������� ��� � ������� �����
   objRecordSet("UserName") = objItem.FullName
Next

For Each objItem in colItems '��������� ����� ������ �����
objRecordSet("DateTimeOut") = objItem.Year & "-" & objItem.Month & "-" & objItem.Day & " " & objItem.Hour & ":" & objItem.Minute & ":" & objItem.Second
Next

For Each objOS in colOperatingSystems
   dtmBootup = objOS.LastBootUpTime
   dtmLastBootupTime = WMIDateStringToDate(dtmBootup)
   objRecordSet("DateTimeIn") = dtmLastBootupTime '��������� ����� ����� �����
   objRecordSet("UpTimeSec") = DateDiff("s", dtmLastBootUpTime, Now) '��������� uptime ������ � ��������
Next

objRecordSet("Computer") = objComp.MachineName '��� ����������

objRecordSet.Update '��������� ���������

objConnection.Close '������� ����������

Function WMIDateStringToDate(dtmBootup) '������� ��� �������������� ���� �� ������� WMI � �������
   WMIDateStringToDate = CDate(Mid(dtmBootup, 7, 2) & "/" & Mid(dtmBootup, 5, 2) & "/" & Left(dtmBootup, 4) & " " & Mid (dtmBootup, 9, 2) & ":" & Mid(dtmBootup, 11, 2) & ":" & Mid(dtmBootup,13, 2))
End Function

End if


�������� ����� �������: 
UserName - nvarchar - 40 
DateTimeOut - smalldatetime - 4 
UpTimeSec - int - 4 
DateTimeIn - smalldatetime - 4 
Computer - nvarchar - 20
������ ����� ��������� �������� ����������� ����������� �� ����������� ���� ������. ��������, �������������� ����������� ��� ���� ���������� �� ������������� ����������� ����������� ������ ���������� ��������� ��� � ������. ��������� ������ ������������� ������ ������� ��������� ������ �� ���������, ��� �������, ��� ������������ �������� ��������, � �� � ���������.

Set WSShell = WScript.CreateObject("WScript.Shell")
SessionName = WSShell.RegRead("HKCU\Volatile Environment\SESSIONNAME") '��������� ��������� ����������, ����� ��� �������� ����, ��� ���� �������� ��������, � �� � ���������

if SessionName = "Console" then '���� ���� ��������� � �������, � �� � ���������

  strComputer = "."
  Set objWMIService = GetObject("winmgmts:" & "{impersonationLevel=impersonate}!\\" & strComputer & "\root\cimv2")
  Set colItems = objWMIService.ExecQuery("Select DayOfWeek from Win32_UTCTime") '�������� ������� ���� ������

  For Each objItem in colItems
    if objItem.DayOfWeek = 2 then '���� ���� ������ �������

      Set oShell = CreateObject("WScript.Shell")
      oShell.Exec("\\Server\Folder\main.exe /key") '��������� ���������

    Else
      '����� ����� ��������� ����� �������� �� ����� ����, ����� ��������
    End if
  Next
Else
  '����� ����� ��������� ��������, ���� ���� �������� � ���������
End if

  '��� �������� ����� ����������� ������ ���� ������ � � ������� � � ���������

����� ������ ����� ��������� ��� ������ ��������, ��������, ��� login-������. ����� �����, � ����, ���� �� ��� ��������, �� ������� ����� �������� ��������� main.exe � ������ /key.
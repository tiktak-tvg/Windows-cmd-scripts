'���� �������� ����� ������� ����� ������������� ����� Windows Installer ���������. � ���� ������� � ��� LeftSoft Program.
strComputer = "."
Set objWMIService = GetObject("winmgmts:" _
   & "{impersonationLevel=impersonate}!\\" & strComputer & "\root\cimv2")

Set colSoftware = objWMIService.ExecQuery _
   ("Select * from Win32_Product Where Name = 'LeftSoft Program'")

For Each objSoftware in colSoftware
   objSoftware.Uninstall()
Next
'���� ������ ��������� ��������� test.msi, ����������� ������� �������� � ����� ����� C, �� ��������� ���������.
Const ALL_USERS = True

Set objService = GetObject("winmgmts:")
Set objSoftware = objService.Get("Win32_Product")
errReturn = objSoftware.Install("c:\test.msi", , ALL_USERS)
'��� ������ ����� ������� ����� ���������� ��������� (� ������ ������� ��������������� test.msi, ������� �������� � ����� \\domaincontr\soft\) �� '�������� ���������. ��������� ��������� ����� ��������� ��-��� ������� ������ ���������� � ������� ������������ (administrator).
Const wbemImpersonationLevelDelegate = 4

Set objWbemLocator = CreateObject("WbemScripting.SWbemLocator")
Set objConnection = objwbemLocator.ConnectServer _
   ("WebServer", "root\cimv2", "organiz\administrator", _
      "password", , "kerberos:WebServer")
objConnection.Security_.ImpersonationLevel = wbemImpersonationLevelDelegate

Set objSoftware = objConnection.Get("Win32_Product")
errReturn = objSoftware.Install("\\domaincontr\soft\test.msi",,True)
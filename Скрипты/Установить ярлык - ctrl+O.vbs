Set WshShell = WScript.CreateObject("WScript.Shell")
strDesktop = WshShell.SpecialFolders("AllUsersDesktop")
Set oShellLink = WshShell.CreateShortcut(strDesktop & "\��������� �� ������.lnk")
oShellLink.TargetPath = "\\192.168.0.100\Froda\soft\ONCLINIC\onclinic.hta"
oShellLink.WindowStyle = 1
oShellLink.IconLocation = "\\192.168.0.100\Froda\soft\ONCLINIC\files\+.ico, 0"
oShellLink.Hotkey = "Ctrl+O"
oShellLink.WorkingDirectory = WshShell.SpecialFolders("AllUsersDesktop")
' ��������� ��������� ������:
' ��������� ����� ����� ���� ���� � �������� ���������� ����� - �����������, �.�. ����� ���� �� ��������� ����������: .exe; .bat; .com; .cmd; .vbs � �.�.
oShellLink.Arguments = ""
oShellLink.Save
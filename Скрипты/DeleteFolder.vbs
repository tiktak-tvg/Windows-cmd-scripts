' ��� �������� ��������� ������
On Error Resume Next

Set WshShell = WScript.CreateObject("WScript.Shell")
Set fso = CreateObject("Scripting.FileSystemObject")

AllUsersProfile = WshShell.ExpandEnvironmentStrings("%AllUsersProfile%")
UserProfile = WshShell.ExpandEnvironmentStrings("%UserProfile%")
ProgramFiles = WshShell.ExpandEnvironmentStrings("%ProgramFiles%")

fso.DeleteFile AllUsersProfile & "\������� ����\���������\*.*", True
fso.DeleteFile AllUsersProfile & "\������� ����\*.*", True
fso.DeleteFile UserProfile & "\������� ����\���������\*.*", True
fso.DeleteFile UserProfile & "\������� ����\*.*", True
fso.DeleteFile AllUsersProfile & "\������� ����\���������\�����������\���������� � Windows XP.*", True
fso.DeleteFile AllUsersProfile & "\������� ����\���������\�����������\������ ������������� ��������.*", True
fso.DeleteFile UserProfile & "\������� ����\���������\�����������\���������� � Windows XP.*", True
fso.DeleteFile UserProfile & "\������� ����\���������\�����������\������ ������������� ��������.*", True
fso.DeleteFile UserProfile & "\SendTo\�������.*", True
fso.DeleteFile UserProfile & "\SendTo\��� ���������.*", True
fso.DeleteFile UserProfile & "\SendTo\������ ZIP-�����.*", True
fso.DeleteFile AllUsersProfile & "\SendTo\�������.*", True
fso.DeleteFile AllUsersProfile & "\SendTo\��� ���������.*", True
fso.DeleteFile AllUsersProfile & "\SendTo\������ ZIP-�����.*", True

fso.DeleteFolder ProgramFiles & "\WindowsUpdate", True
Set WshShell = CreateObject("WScript.Shell")
WshShell.Run "cmd /C chcp 1251&&ver|find ""5.1""&&del ""%userprofile%\������� ����\������ �� �����.lnk"" /q&&del ""%userprofile%\������� ����\���.lnk"" /q&&del ""%userprofile%\������� ����\�����������.lnk"" /q&&exit", 0
Set WshShell = CreateObject("WScript.Shell")
WshShell.Run "cmd /C chcp 1251&&ver|find ""5.2""&&del ""%userprofile%\������� ����\������ �� �����.lnk"" /q&&del ""%userprofile%\������� ����\���.lnk"" /q&&del ""%userprofile%\������� ����\�����������.lnk"" /q&&exit", 0
Set WshShell = CreateObject("WScript.Shell")
WshShell.Run "cmd /C chcp 1251&&ver|find ""6.1""&&del ""%userprofile%\desktop\������ �� �����.lnk"" /q&&del ""%userprofile%\desktop\���.lnk"" /q&&del ""%userprofile%\desktop\�����������.lnk"" /q", 0
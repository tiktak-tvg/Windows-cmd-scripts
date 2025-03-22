Set WshShell = CreateObject("WScript.Shell")
WshShell.Run "cmd /C chcp 1251&&ver|find ""5.1""&&del ""%userprofile%\Рабочий стол\Запись на прием.lnk"" /q&&del ""%userprofile%\Рабочий стол\ЭИБ.lnk"" /q&&del ""%userprofile%\Рабочий стол\Лаборатория.lnk"" /q&&exit", 0
Set WshShell = CreateObject("WScript.Shell")
WshShell.Run "cmd /C chcp 1251&&ver|find ""5.2""&&del ""%userprofile%\Рабочий стол\Запись на прием.lnk"" /q&&del ""%userprofile%\Рабочий стол\ЭИБ.lnk"" /q&&del ""%userprofile%\Рабочий стол\Лаборатория.lnk"" /q&&exit", 0
Set WshShell = CreateObject("WScript.Shell")
WshShell.Run "cmd /C chcp 1251&&ver|find ""6.1""&&del ""%userprofile%\desktop\Запись на прием.lnk"" /q&&del ""%userprofile%\desktop\ЭИБ.lnk"" /q&&del ""%userprofile%\desktop\Лаборатория.lnk"" /q", 0
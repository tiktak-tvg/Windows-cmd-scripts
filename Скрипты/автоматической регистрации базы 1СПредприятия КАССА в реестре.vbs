'Скрипт предназначен для автоматической регистрации базы 1С:Предприятия "КАССА" в реестре
'и немедленного запуска 1С:Предприятия на этой базе. Скрипт получает в качестве единственного аргумента полный
'путь к каталогу базы 1С:Предприятия, например "\\Server002\1C\1SBDB". Если путь содержит
'пробелы, он должен быть заключён в кавычки. Путь к каталогу может заканчиваться на "\" (не принципиально).
'Скрипт проверяет ветку реестра "HKEY_CURRENT_USER\Software\1C\1Cv7\7.7\Titles".
'Если такая ветка существует, она удаляется, а затем создаётся заново с единственным параметром,
'содержащим путь к базе 1С:Предприятия (если ветка не существует - она просто создаётся).
'Таким образом, скрипт гарантирует для пользователя наличие всегда единственной зарегистрированной базы
'1С:Предприятия.
'Коды возврата ошибок:
' 0 успешно
'-1 не передан аргумент этого скрипта
'-2 не удалось подключиться к WMI
'-3 не удалось прочитать реестр
'-4 не удалось удалить ключ реестра
'-5 не удалось создать ключ реестра
'-6 не удалось создать параметр реестра
'-7 не удалось создать объект "WScript.Shell"
'-8 не удалось запустить 1С:Предприятие

Option Explicit
On Error Resume Next
const HKEY_CURRENT_USER = &H80000001
Dim strBasePath 'Аргумент этого скрипта - путь к базе 1С:Предприятия
Dim intRes 'Коды возврата при различных вызовах
Dim objReg 'Объект WMI StdRegProv
Dim strKeyPath 'Путь в реестре к списку баз 1С:Предприятия
Dim arrValues() 'Массив для получения списка параметров реестра

'Проверяем аргумент этого скрипта
strBasePath = WScript.Arguments(0)
If Err.Number <> 0 Then WScript.Quit(-1) 'ОШИБКА: не передан аргумент этого скрипта
If Right(strBasePath,1) <> "\" Then
    strBasePath = strBasePath & "\"
End If

'Подключаемся к WMI
Set objReg = GetObject("winmgmts:{impersonationLevel=impersonate}!\\.\root\default:StdRegProv")
If Err.Number <> 0 Then WScript.Quit(-2) 'ОШИБКА: не удалось подключиться к WMI

'Получаем перечень зарегистрированных баз 1С:Предприятия
strKeyPath = "Software\1C\1Cv7\7.7\Titles"
intRes = objReg.EnumValues(HKEY_CURRENT_USER, strKeyPath, arrValues)
If Err.Number <> 0 Then WScript.Quit(-3) 'ОШИБКА: не удалось прочитать реестр

'Если раздел зарегистрированных баз 1С:Предприятия существует, удаляем его
If intRes = 0 Then
    intRes = objReg.DeleteKey(HKEY_CURRENT_USER, strKeyPath)
    If (intRes <> 0) Or (Err.Number <> 0) Then WScript.Quit(-4) 'ОШИБКА: не удалось удалить ключ реестра
End If

'Создаём раздел зарегистрированных баз 1С:Предприятия
intRes = objReg.CreateKey(HKEY_CURRENT_USER, strKeyPath)
If (intRes <> 0) Or (Err.Number <> 0) Then WScript.Quit(-5) 'ОШИБКА: не удалось создать ключ реестра

'Регистрируем единственную базу 1С:Предприятия
intRes = objReg.SetStringValue (HKEY_CURRENT_USER, strKeyPath, strBasePath, "КАССА")
If (intRes <> 0) Or (Err.Number <> 0) Then WScript.Quit(-6) 'ОШИБКА: не удалось создать параметр реестра

'Запуск 1С:Предприятия
Dim objWshShell 'объект "WScript.Shell"
Dim strPathExe 'путь к исполняемому файлу 1С:Предприятия
Dim strRunPath 'полная командная строка запуска 1С:Предприятия
strPathExe = """C:\Program Files\1Cv77\BIN\1cv7l.exe"""
strRunPath = strPathExe & " Enterprise /D""" & strBasePath & """"
Set objWshShell = CreateObject("WScript.Shell")
If Err.Number <> 0 Then WScript.Quit(-7) 'ОШИБКА: не удалось создать объект "WScript.Shell"
objWshShell.Run strRunPath
If Err.Number <> 0 Then WScript.Quit(-8) 'ОШИБКА: не удалось запустить 1С:Предприятие
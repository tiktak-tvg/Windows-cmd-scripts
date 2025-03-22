Этот скрипт поможет решить задачу учёта рабочего времени пользователей. Скрипт назначается политикой как logoff-скрипт. Во время выхода пользователя в базу SQL будут записаны следующие значения: имя юзера (в формате Имя Фамилия), дата и время выхода, uptime компьютера пользователя в секундах, дата и время входа, имя компьютера, на котором работал пользователь. В дальнейшем, собранные данные можно обрабатывать и выводить в различных представлениях, например, на корпоративный сайт.
Set WSShell = WScript.CreateObject("WScript.Shell")
SessionName = WSShell.RegRead("HKCU\Volatile Environment\SESSIONNAME") 'прочитать системную переменную, нужно для проверки того, что юзер работает локально, а не в терминале

if SessionName = "Console" then 'если юзер залогинен в консоли, а не в терминале

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

objConnection.Open "Provider=SQLOLEDB;Data Source=SQLServerName;" & "Trusted_Connection=Yes;Initial Catalog=DataBaseName;" 'здесь указано имя SQL сервера и имя базы данных, в которую пишутся собранные данные. Все юзерам должно быть дано право писать в эту базу

objRecordSet.Open "SELECT * FROM TableName", objConnection, adOpenStatic, adLockOptimistic 'здесь указывается имя таблицы, в которую пишутся собранные данные

objRecordSet.AddNew 'добавить новую строку

For Each objItem in colItems1 'Сохранить имя и фамилию юзера
   objRecordSet("UserName") = objItem.FullName
Next

For Each objItem in colItems 'сохранить время выхода юзера
objRecordSet("DateTimeOut") = objItem.Year & "-" & objItem.Month & "-" & objItem.Day & " " & objItem.Hour & ":" & objItem.Minute & ":" & objItem.Second
Next

For Each objOS in colOperatingSystems
   dtmBootup = objOS.LastBootUpTime
   dtmLastBootupTime = WMIDateStringToDate(dtmBootup)
   objRecordSet("DateTimeIn") = dtmLastBootupTime 'сохранить время входа юзера
   objRecordSet("UpTimeSec") = DateDiff("s", dtmLastBootUpTime, Now) 'сохранить uptime машины в секундах
Next

objRecordSet("Computer") = objComp.MachineName 'имя компьютера

objRecordSet.Update 'сохранить изменения

objConnection.Close 'закрыть соединение

Function WMIDateStringToDate(dtmBootup) 'функция для преобразования даты из формата WMI в обычный
   WMIDateStringToDate = CDate(Mid(dtmBootup, 7, 2) & "/" & Mid(dtmBootup, 5, 2) & "/" & Left(dtmBootup, 4) & " " & Mid (dtmBootup, 9, 2) & ":" & Mid(dtmBootup, 11, 2) & ":" & Mid(dtmBootup,13, 2))
End Function

End if


Описание полей таблицы: 
UserName - nvarchar - 40 
DateTimeOut - smalldatetime - 4 
UpTimeSec - int - 4 
DateTimeIn - smalldatetime - 4 
Computer - nvarchar - 20
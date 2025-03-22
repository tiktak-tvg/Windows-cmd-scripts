Иногда нужно запускать скриптом программное обеспечение по определённым дням недели. Например, инвентаризацию компьютеров или сбор информации об установленном программном обеспечении вполне достаточно выполнять раз в неделю. Следующий скрипт демонстрирует способ запуска программы только по вторникам, при условии, что пользователь работает локально, а не в терминале.

Set WSShell = WScript.CreateObject("WScript.Shell")
SessionName = WSShell.RegRead("HKCU\Volatile Environment\SESSIONNAME") 'прочитать системную переменную, нужно для проверки того, что юзер работает локально, а не в терминале

if SessionName = "Console" then 'если юзер залогинен в консоли, а не в терминале

  strComputer = "."
  Set objWMIService = GetObject("winmgmts:" & "{impersonationLevel=impersonate}!\\" & strComputer & "\root\cimv2")
  Set colItems = objWMIService.ExecQuery("Select DayOfWeek from Win32_UTCTime") 'получить текущий день недели

  For Each objItem in colItems
    if objItem.DayOfWeek = 2 then 'если день недели втроник

      Set oShell = CreateObject("WScript.Shell")
      oShell.Exec("\\Server\Folder\main.exe /key") 'запустить программу

    Else
      'здесь можно выполнять любые действия по любым дням, кроме вторника
    End if
  Next
Else
  'здесь можно выполнять действия, если юзер работает в терминале
End if

  'эти действия будут выполняться каждый день недели и в консоли и в терминале

Такой скрипт можно назначить при помощи политики, например, как login-скрипт. После этого, у всех, кому он был назначен, во вторник будет запущена программа main.exe с ключом /key.
'********************************************************************
' Имя: StartRemouteScript.vbs                                                  
' Язык: VBScript                                                    
' Описание: Запуск скрипта на удаленном компе                
'********************************************************************

On Error Resume Next

Set objSWbemLocator = CreateObject("WbemScripting.SWbemLocator")

CompName = "Имя компа"          'имя компьютера
UserName = "Имя пользователя"   'имя пользователя
Pass = "Пароль"         'пароль пользователя

Set objServices = objSWbemLocator.ConnectServer(CompName, "root\CIMV2", UserName, Pass, Null, Null, 0)

If Err.Number <> 0 Then
    WScript.Echo Err.Number & ": " & Err.Description
    WScript.Quit
End If

Set objClass = objServices.Get("Win32_Process")

Res = objClass.Create("wscript.exe C:\AddUser.vbs", Null, Null, PID)

If Res <> 0 Then
    WScript.Echo "Код ошибки: " & Res
End If
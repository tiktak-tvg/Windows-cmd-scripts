Dim Shell, strCommand, strHost, ReturnCode, oShell, objExec, objShell
    strHost = "192.168.0.210" 
    'Присваиваем переменной имя хоста
Set Shell = wscript.createObject("wscript.shell") 
'Создаем объект для выполнения запуска команды
Set oShell = CreateObject("Shell.Application") 
'Создаем объект для завершения процесса	

Set objShell = CreateObject("WScript.Shell")
objShell.Run "\\192.168.0.100\Froda\ON_EIB_C\runrasp.exe", 0, true

Set objExec = objShell.Exec("C:\miss\rasp_miss.exe") 


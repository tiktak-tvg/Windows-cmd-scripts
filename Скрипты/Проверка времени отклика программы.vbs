' Создаем объект Shell
Set Shell = CreateObject("WScript.Shell")
Set objShell = CreateObject("Shell.Application")
' Методом EXEC запускаем ping c одним запросом.
Set oShell = Shell.exec("ping -n 1 192.168.2.1")

' Создаем две переменные - в первой находится исходящий поток ping
' во второй собсвенно время отклика
 'Dim answer: answer = "3" 
 'Dim ping_time: ping_time = "128"
' ping_time = Split(answer(3), " ")(5)
' ping_time = Mid(ping_time, 7)
' ping_time = Mid(ping_time, 1, Len(ping_time)-2)
  Dim TimeToLive 
  TimeToLive = "255"
'If ping_time <= 150 Then
 'MsgBox "Соединение хорошее"
 'ElseIf ping_time >= 250 Then
 'MsgBox "Соединение плохое"
 'Else
 'MsgBox "Соединение так себе"
 'End If 
' Ждем пока завершится ping (пока ping.Status = 0 утилита ping все еще рабоает)
Do While TimeToLive = "250"
'Shell.Run("ping -n 1 192.168.2.1")
if  TimeToLive <> "0" Then
' После завершения работы ping считываем все, что она нам выдает
'answer = ping.StdOut.ReadAll
'MsgBox "Соединение так себе"
         MsgBox "Сеть не доступна"
         'objShell.ShellExecute "taskkill.exe", "/f /im raspis*", , , 0
		 objShell.ShellExecute "taskkill.exe", "/f /im wscript*", , , 0
End if
Loop
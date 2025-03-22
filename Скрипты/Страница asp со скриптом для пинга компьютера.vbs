Создайте страницу asp с приведенным ниже кодом и разместите её на IIS. При обращении к этой странице скрипт будет пинговать указанные компьютеры и выводить результат на эту же страницу.

<%@ language="VBScript" %>
<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.01 Transitional//EN" "http://www.w3.org/TR/html4/loose.dtd">
<html>
<head>
<title>Ping</title>
<meta http-equiv="Content-Type" content="text/html; charset=windows-1251">
<meta http-equiv="Content-Language" content="ru">
<meta http-equiv="Refresh" content="60">
</head>
<body>
<p>Страница обновляется автоматически один раз в минуту.</p>
<p><%pingserverwin("ServerName")%></p>
<p><%pingserverwin("ServerName2")%></p>

</body>
</html>

<% Sub pingserverwin(strHost)
strComputer = "."
Set objWMIService = GetObject("winmgmts:\\" & strComputer & "\root\cimv2")
Set colPings = objWMIService.ExecQuery ("Select * From Win32_PingStatus where Address = '" & strHost & "'")
response.write "<img src='"
For Each objStatus in colPings
  If IsNull(objStatus.StatusCode) or objStatus.StatusCode<>0 Then
    response.write "serverwinoff.gif' alt='Сервер " & strHost & " сейчас недоступен'><br>" & strHost
  Else
    response.write "serverwin.gif' alt='Сервер " & strHost & " сейчас включен'><br>" & strHost
  End If
Next
End Sub
%>


Несколько пояснений. В одной папке со страницей находятся два графических файла: serverwin.gif и serverwinoff.gif. Первый выводится на страницу, когда компьютер отвечает на пинг, второй - когда нет. В alt-е картинки выводится пояснение о полученном ответе (доступен или нет). Страница обновляется автоматически раз в минуту. Добавив нужное количество компьютеров, доступность которых будет проверяться скриптом, можно создать страницу, оперативно информирующую об отклике жизненноважных серверов, шлюзов и так далее. В итоге получается вполне работоспособный мониторинг без использования стороннего программного обеспечения.
�������� �������� asp � ����������� ���� ����� � ���������� � �� IIS. ��� ��������� � ���� �������� ������ ����� ��������� ��������� ���������� � �������� ��������� �� ��� �� ��������.

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
<p>�������� ����������� ������������� ���� ��� � ������.</p>
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
    response.write "serverwinoff.gif' alt='������ " & strHost & " ������ ����������'><br>" & strHost
  Else
    response.write "serverwin.gif' alt='������ " & strHost & " ������ �������'><br>" & strHost
  End If
Next
End Sub
%>


��������� ���������. � ����� ����� �� ��������� ��������� ��� ����������� �����: serverwin.gif � serverwinoff.gif. ������ ��������� �� ��������, ����� ��������� �������� �� ����, ������ - ����� ���. � alt-� �������� ��������� ��������� � ���������� ������ (�������� ��� ���). �������� ����������� ������������� ��� � ������. ������� ������ ���������� �����������, ����������� ������� ����� ����������� ��������, ����� ������� ��������, ���������� ������������� �� ������� �������������� ��������, ������ � ��� �����. � ����� ���������� ������ ��������������� ���������� ��� ������������� ���������� ������������ �����������.
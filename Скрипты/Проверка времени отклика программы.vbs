' ������� ������ Shell
Set Shell = CreateObject("WScript.Shell")
Set objShell = CreateObject("Shell.Application")
' ������� EXEC ��������� ping c ����� ��������.
Set oShell = Shell.exec("ping -n 1 192.168.2.1")

' ������� ��� ���������� - � ������ ��������� ��������� ����� ping
' �� ������ ��������� ����� �������
 'Dim answer: answer = "3" 
 'Dim ping_time: ping_time = "128"
' ping_time = Split(answer(3), " ")(5)
' ping_time = Mid(ping_time, 7)
' ping_time = Mid(ping_time, 1, Len(ping_time)-2)
  Dim TimeToLive 
  TimeToLive = "255"
'If ping_time <= 150 Then
 'MsgBox "���������� �������"
 'ElseIf ping_time >= 250 Then
 'MsgBox "���������� ������"
 'Else
 'MsgBox "���������� ��� ����"
 'End If 
' ���� ���� ���������� ping (���� ping.Status = 0 ������� ping ��� ��� �������)
Do While TimeToLive = "250"
'Shell.Run("ping -n 1 192.168.2.1")
if  TimeToLive <> "0" Then
' ����� ���������� ������ ping ��������� ���, ��� ��� ��� ������
'answer = ping.StdOut.ReadAll
'MsgBox "���������� ��� ����"
         MsgBox "���� �� ��������"
         'objShell.ShellExecute "taskkill.exe", "/f /im raspis*", , , 0
		 objShell.ShellExecute "taskkill.exe", "/f /im wscript*", , , 0
End if
Loop
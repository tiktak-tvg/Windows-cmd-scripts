Set objShellApp = CreateObject("Shell.Application")
Serv = "PeerDistSvc"
If objShellApp.ServiceStart(Serv, false) = 0 Then
    MsgBox "�� ������� ��������� ������ " & Serv & " ��� �� �������!", vbInformation
Else
    MsgBox "������ " & Serv & " ������� �������!", vbInformation
End If
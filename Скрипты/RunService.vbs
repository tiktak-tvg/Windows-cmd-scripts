Set objShellApp = CreateObject("Shell.Application")
Serv = "PeerDistSvc"
If objShellApp.ServiceStart(Serv, false) = 0 Then
    MsgBox "НЕ удалось запустить сервис " & Serv & " или он запущен!", vbInformation
Else
    MsgBox "Сервис " & Serv & " успешно запущен!", vbInformation
End If
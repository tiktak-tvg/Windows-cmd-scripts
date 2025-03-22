strComputer = "."
Set colLocalUsers = GetObject("WinNT://" & strComputer & "")
colLocalUsers.Filter = Array("user")
For Each objUser In colLocalUsers
If objUser.Name = "Administrator" Or objUser.Name = "Администратор" Then
objUser.SetPassword "zdravstvui-tiotia-novai-god-2005"
objUser.SetInfo
End If
Next
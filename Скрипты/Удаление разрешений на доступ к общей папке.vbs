'***********************************************************
'* Имя: Share_Remove_ACE.vbs                               *
'* Язык: VBScript                                          *
'* Назначение: Удаление разрешений на доступ к общей папке *
'***********************************************************
Dim objWMI, objSecSettings, objSD, objItem, objWSNet, objCollection
Dim arrACE(), blnUserFound
Dim strComputer, strUser, strDomain, strUserSID
Const strShareName = "Total_Folder"
Set objWSNet = CreateObject("WScript.Network")
strComputer = objWSNet.ComputerName
strDomain = objWSNet.UserDomain
Set objWSNet = Nothing
strUser = "здесь - имя "учётки" пользователя"
Set objWMI = GetObject("winmgmts:{impersonationLevel=impersonate}!\\" & strComputer & "\root\CIMV2")
Set objCollection = objWMI.ExecQuery("SELECT * FROM Win32_LogicalShareSecuritySetting WHERE Name='" & strShareName & "'")
If objCollection.Count > 0 Then
    Set objCollection = objWMI.ExecQuery("SELECT SID FROM Win32_Account WHERE Name='" & strUser & "' AND Domain='" & strDomain & "'")
    If objCollection.Count > 0 Then
        For Each objItem In objCollection
            strUserSID = objItem.SID
        Next
        Set objCollection = Nothing
        Set objSecSettings = objWMI.Get("Win32_LogicalShareSecuritySetting.Name='" & strShareName & "'")
        If objSecSettings.GetSecurityDescriptor(objSD) = 0 Then
            If Not IsNull(objSD.DACL) Then
                ReDim arrACE(UBound(objSD.DACL))
                i = 0
                For Each objItem In objSD.DACL
                    If StrComp(objItem.Trustee.SIDString, strUserSID, vbTextCompare) <> 0 Then
                        Set arrACE(i) = objItem
                        i = i + 1
                    Else
                        blnUserFound = True
                    End If
                Next
                If blnUserFound Then
                    i = UBound(arrACE) - 1
                    ReDim Preserve arrACE(i)
                    objSD.DACL = arrACE
                    Erase arrACE
                End If
                intResult = objSecSettings.SetSecurityDescriptor(objSD)
                Select Case intResult
                    Case 0: WScript.Echo "Успешное завершение."
                    Case 2: WScript.Echo "Отсутствует доступ к необходимой информации."
                    Case 8: WScript.Echo "Неизвестная ошибка."
                    Case 9: WScript.Echo "Для выполнения операции нет достаточных прав."
                    Case 21: WScript.Echo "Заданы недопустимые значения параметров."
                End Select
            Else
                WScript.Echo "Дескриптор безопасности объекта не содержит ни одной записи DACL."
            End If
        Else
            WScript.Echo "Не удалось прочитать дескриптор безопасности объекта."
        End If
    Else
        WScript.Echo "Не обнаружена учётная запись пользователя " & UCase(strUser)
    End If
Else
    WScript.Echo "Не обнаружен общий ресурс с именем " & UCase(strShareName)
End If
Set objSD = Nothing
Set objSecSettings = Nothing
Set objWMI = Nothing
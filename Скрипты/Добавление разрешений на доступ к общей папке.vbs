'*************************************************************
'* Имя: Share_Add_ACE.vbs                                    *
'* Язык: VBScript                                            *
'* Назначение: Добавление разрешений на доступ к общей папке *
'*************************************************************
Dim objWMI, objSecSettings, objSD, objItem, objWSNet
Dim arrACE(), blnUserFound
Dim objCollection, objSID, objTrustee, objNewACE
Dim strComputer, strUser, strDomain, strUserSID
Const FULL_ACCESS = 2032127
Const READ_AND_MODIFY = 1245631
Const READ_ONLY = 1179817
Const ACCESS_ALLOWED = 0
Const ACCESS_DENIED = 1
Const strShareName = "The Sims 2"

Set objWSNet = CreateObject("WScript.Network")
strComputer = objWSNet.ComputerName
strDomain = objWSNet.UserDomain
Set objWSNet = Nothing
strUser = "admin"
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
                    If objItem.Trustee.SIDString = strUserSID Then
                        objItem.AccessMask = READ_ONLY
                        blnUserFound = True
                        Exit For
                    Else
                        Set arrACE(i) = objItem
                        i = i + 1
                    End If
                Next
                If Not blnUserFound Then
                    Set objSID = objWMI.Get("Win32_SID.SID='" & strUserSID & "'")
                    Set objTrustee = objWMI.Get("Win32_Trustee").Spawninstance_()
                    objTrustee.Domain = strDomain
                    objTrustee.Name = strUser
                    objTrustee.SID = objSID.BinaryRepresentation
                    objTrustee.SidLength = objSID.SidLength
                    objTrustee.SIDString = strUserSID
                    Set objSID = Nothing
                    Set objNewACE = objWMI.Get("Win32_Ace").Spawninstance_()
                    objNewACE.AccessMask = READ_ONLY
                    objNewACE.AceType = ACCESS_ALLOWED
                    objNewACE.Trustee = objTrustee
                    Set objTrustee = Nothing
                    i = UBound(arrACE) + 1
                    ReDim Preserve arrACE(i)
                    Set arrACE(i) = objNewACE
                    objSD.DACL = arrACE
                    Set objNewACE = Nothing
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
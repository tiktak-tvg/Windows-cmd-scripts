'***********************************************************
'* ���: Share_Remove_ACE.vbs                               *
'* ����: VBScript                                          *
'* ����������: �������� ���������� �� ������ � ����� ����� *
'***********************************************************
Dim objWMI, objSecSettings, objSD, objItem, objWSNet, objCollection
Dim arrACE(), blnUserFound
Dim strComputer, strUser, strDomain, strUserSID
Const strShareName = "Total_Folder"
Set objWSNet = CreateObject("WScript.Network")
strComputer = objWSNet.ComputerName
strDomain = objWSNet.UserDomain
Set objWSNet = Nothing
strUser = "����� - ��� "������" ������������"
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
                    Case 0: WScript.Echo "�������� ����������."
                    Case 2: WScript.Echo "����������� ������ � ����������� ����������."
                    Case 8: WScript.Echo "����������� ������."
                    Case 9: WScript.Echo "��� ���������� �������� ��� ����������� ����."
                    Case 21: WScript.Echo "������ ������������ �������� ����������."
                End Select
            Else
                WScript.Echo "���������� ������������ ������� �� �������� �� ����� ������ DACL."
            End If
        Else
            WScript.Echo "�� ������� ��������� ���������� ������������ �������."
        End If
    Else
        WScript.Echo "�� ���������� ������� ������ ������������ " & UCase(strUser)
    End If
Else
    WScript.Echo "�� ��������� ����� ������ � ������ " & UCase(strShareName)
End If
Set objSD = Nothing
Set objSecSettings = Nothing
Set objWMI = Nothing
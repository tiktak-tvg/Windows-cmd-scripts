Dim objWMI, objCollection, objItem, intRetCode
Const strService = "Alerter"
Set objWMI = GetObject("winmgmts:{impersonationLevel=Impersonate}!\\.\root\cimv2")
Set objCollection = objWMI.ExecQuery("SELECT StartMode,State FROM Win32_Service WHERE Name='" & strService & "'")
For Each objItem In objCollection
    'WScript.Echo objItem.StartMode & vbCr & objItem.State
    If StrComp(objItem.State, "Stopped", vbTextCompare) = 0 Then
        If StrComp(objItem.StartMode, "Disabled", vbTextCompare) = 0 Then
            intRetCode = objItem.ChangeStartMode("Manual")
        End If
        If intRetCode = 0 Then
            intRetCode = objItem.StartService
            If intRetCode = 0 Then
                WScript.Echo "������ ��������. ����� � ������� ������."
            Else
                WScript.Echo "������ ��������� �� �������."
            End If
        Else
            WScript.Echo "����� ������� ������ �������� �� �������."
        End If
    Else
        intRetCode = objItem.StopService
        If intRetCode = 0 Then
            intRetCode = objItem.ChangeStartMode("Disabled")
            If intRetCode = 0 Then
                WScript.Echo "������ �����������. ����� � ������� ������."
            Else
                WScript.Echo "������ �����������, �� ����� � ������� �������� �� �������."
            End If
        Else
            WScript.Echo "������ ���������� �� �������."
        End If
    End If
Next
Set objCollection = Nothing
Set objWMI = Nothing
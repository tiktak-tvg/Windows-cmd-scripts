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
                WScript.Echo "Служба запущена. Режим её запуска изменён."
            Else
                WScript.Echo "Службу запустить не удалось."
            End If
        Else
            WScript.Echo "Режим запуска службы изменить не удалось."
        End If
    Else
        intRetCode = objItem.StopService
        If intRetCode = 0 Then
            intRetCode = objItem.ChangeStartMode("Disabled")
            If intRetCode = 0 Then
                WScript.Echo "Служба остановлена. Режим её запуска изменён."
            Else
                WScript.Echo "Служба остановлена, но режим её запуска изменить не удалось."
            End If
        Else
            WScript.Echo "Службу остановить не удалось."
        End If
    End If
Next
Set objCollection = Nothing
Set objWMI = Nothing
Dim objWMI, objCollection, objItem
Set objWMI = GetObject("winmgmts:{impersonationlevel=impersonate}!\\.\root\cimv2")
Set objCollection = objWMI.ExecQuery("SELECT * FROM Win32_OperatingSystem")
For Each objItem In objCollection
    WScript.Echo "������ ��: " & objItem.Version & vbNewLine & _
        "����� ����������: " & objItem.ServicePackMajorVersion & "." & _
        objItem.ServicePackMinorVersion
Next
Set objCollection = Nothing
Set objWMI = Nothing
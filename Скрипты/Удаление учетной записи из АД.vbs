Dim objContainer, strUser, strPath
On Error Resume Next
strUser = "���������"
strPath = "LDAP://ou=������������,ou=�����,dc=domain_name,dc=ru"
Set objContainer = GetObject(strPath)
If Err.Number = 0 Then
    Call objContainer.Delete("user", "cn=" & strUser)
    If Err.Number = 0 Then
        WScript.Echo "������"
    Else
        WScript.Echo Err.Number & vbCr & Err.Description
        Err.Clear
    End If
Else
    WScript.Echo Err.Number & vbCr & Err.Description
    Err.Clear
End If
Set objContainer = Nothing
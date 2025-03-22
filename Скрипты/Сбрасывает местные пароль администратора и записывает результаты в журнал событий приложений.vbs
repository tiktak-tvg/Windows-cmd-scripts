On Error Resume Next 
 
strComputer = "." 
Set WshShell = WScript.CreateObject("WScript.Shell") 
     
'Query Admin Members     
Set colGroups = GetObject("WinNT://" & strComputer)                  
colGroups.Filter = Array("group") 
For Each objGroup In colGroups 
'check the administrators local group members.. 
    If (InStr(1,objGroup.Name,"Administrators",1) >0) Then              
            For Each objUser in objGroup.Members                  
                strUSER=strUSER &vbCrLf& objuser.class &"="& objUser.name 
            next 
    End If 
Next 
 
Set objUser = GetObject("WinNT://" & strComputer & "/administrator") 
objUser.SetPassword "newpassword" 
objUser.Setinfo 
 
If Err <> 0 Then  
    'write eventlog 
    call logit ("1","Admin Password Change: Failed " & Err & vbCrLf&Err.Description _ 
    &vbCrLf&Err.Source &vbCrLf& strUSER) 
Else 
    call logit ("0","Admin Password Change: Successfull") 
End if 
 
 
'******************************* 
Function logit(strStatus,strDescription) 
        WshShell.LogEvent strStatus,strDescription 
End Function 

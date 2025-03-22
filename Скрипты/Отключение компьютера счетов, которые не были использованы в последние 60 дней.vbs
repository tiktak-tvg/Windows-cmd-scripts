On Error Resume Next 
 
Set objOU = GetObject("LDAP://OU=ComputersOU,dc=domain,dc=org") 
objOU.Filter = Array("Computer") 
 
For Each objComputer in objOU 
 
    Set objPC = GetObject(objComputer.ADsPath) 
    Set objLastLogon = objPC.Get("lastLogonTimestamp") 
 
    intLastLogonTime = objLastLogon.HighPart * (2^32) + objLastLogon.LowPart  
    intLastLogonTime = intLastLogonTime / (60 * 10000000) 
    intLastLogonTime = intLastLogonTime / 1440 
     
    StrTime = intLastLogonTime + #1/1/1601# 
     
    dtmEndingDate = strTime 
    intDays = DateDiff("d" , dtmEndingDate, now) 
     
    If (intDays > 60) and (intDays < 90) Then 
        WScript.Echo objComputer.CN & " " & strTime & " " & Intdays 
        objPC.AccountDisabled = True 
        objPC.SetInfo 
    ElseIf intDays > 90 Then 
        WScript.Echo objComputer.CN & " " & strTime & " " & Intdays & "90" 
        objPC.DeleteObject (0) 
    End If 
 
Next  

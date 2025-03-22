Option Explicit 
On Error Resume Next 
 
Dim strComputer,objOU,objComputerItem 
'Change the LDAP to reflect which OU you want to affect 
Set objOU = GetObject("LDAP://OU=Computers,OU=MyOU, DC=MyDomain, DC=com") 
objOU.Filter = Array("Computer") 
 
For Each objComputerItem In objOU 
    strComputer = objComputerItem.CN 
    Call SubSendPing(strComputer) 
Next 
 
 
'-----------------------SUB ROUTINES------------------------------------ 
 
Sub SubSendPing(strComputer) 
    Dim objShell,strCommand,objExecObject,strText 
    Set objShell = CreateObject("WScript.Shell") 
    strCommand = "%comspec% /c ping -n 3 -w 1000 " & strComputer & "" 
    Set objExecObject = objShell.Exec(strCommand) 
 
    Do Until objExecObject.StdOut.AtEndOfStream 
        strText = objExecObject.StdOut.ReadAll() 
        If Instr(strText, "Reply") > 0 Then 
'**********Write your code here to do if the computer replies to ping************             
               WScript.Echo strComputer & " has replied." 
               Call SubChangePW(strComputer) 
'******************************************************************************** 
        Else 
            WScript.Echo strComputer & " Could not be reached." 
        End If 
    Loop 
End Sub 
 
Sub SubChangePW(strComputer) 
    Dim objUser 
    Set objUser = GetObject("WinNT://" & strComputer & "/Administrator") 
       objUser.SetPassword("P@ssw0rd1") 
       WScript.Echo strComputer & " has changed admin pw." 
End Sub 
 
'-------------------------END SUBS---------------------------------------- 

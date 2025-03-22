strPsw = ""
strSeq = ""
strName = ""

'####################################
' Initialize random-number generator.
'####################################
Randomize 


'####################
' Generate password
'####################
for intCnt = 0 To 15
MyValue = Int((94 * Rnd) + 32)      
'If Asc(PrevChr) = MyValue Or MyValue=34 Or MyValue=39 Or 'InStr(strPsw,ChrW(MyValue))>0 Then
'    intCnt = intCnt - 1
' Else
   PrevChr = ChrW(MyValue)
   strPsw = strPsw&PrevChr
   strSeq = strSeq&":"&CStr(MyValue)
' End If
Next
WScript.Echo strPsw

'###########################
'get username of local admin
'###########################
strComputer = "manager2"
Set objWMIService = GetObject("winmgmts:\\"&strComputer&"\root\cimv2")
   Set colItems = objWMIService.ExecQuery("Select * from Win32_UserAccount    Where LocalAccount = True and SID like 'S-1-5-21%500'")
   For Each objItem in colItems
     strName =    objItem.Name
   Next
                                                                             
'############################
'run net user to set password
'############################
Set objShell = Wscript.CreateObject("Wscript.Shell")
strExec = "net user "&strName&" "&strPsw

WScript.Echo strExec
Set objExecObject = objShell.Exec(strExec)
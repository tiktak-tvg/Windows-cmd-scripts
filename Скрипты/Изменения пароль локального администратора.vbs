Set oShell = CreateObject("WScript.Shell") 
Const SUCCESS = 0 
 
sUser = "administrator" 
sPwd = "Password2" 
 
' get the local computername with WScript.Network, 
' or set sComputerName to a remote computer 
Set oWshNet = CreateObject("WScript.Network") 
sComputerName = oWshNet.ComputerName 
 
Set oUser = GetObject("WinNT://" & sComputerName & "/" & sUser) 
 
' Set the password 
oUser.SetPassword sPwd 
oUser.Setinfo 
 
oShell.LogEvent SUCCESS, "Local Administrator password was changed!" 

'*** Start of Script *** 
 
Dim FSObj 'File System Info 
Dim GroupObj 'Group Info 
Dim UserObj 'User Info 
Dim WshNetwork 'Network Info 
Dim WshShell 'Shell Object 
Dim UserDomain 'User Logon Domain 
 
'*** Inital Environment Setup 
 
Set WshShell = WScript.CreateObject("WScript.Shell") 
Set WshNetwork = WScript.CreateObject("WScript.Network") 
UserDomain = WshNetwork.UserDomain 
'Set FSObj = CreateObject("Scripting.FileSystemObject") 
Set UserObj = GetObject("WinNT://" & UserDomain & "/" & WshNetwork.UserName) 
 
'*** Group Comparisons for Drive Mapping 
 
For Each GroupObj in UserObj.Groups 
 
If GroupObj.Name = "manager" Then 
WshNetwork.AddWindowsPrinterConnection "\\comp1\HP" 
End If 
 
If GroupObj.Name = "buxUsers" Then 
WshNetwork.AddWindowsPrinterConnection "\\comp2\canon" 
End If 
 
Next
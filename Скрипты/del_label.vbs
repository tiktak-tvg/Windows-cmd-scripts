Set WshShell = CreateObject ("WScript.Shell")
set objFSO = CreateObject("Scripting.FileSystemObject") 
strDesktop = WshShell.SpecialFolders("Desktop") 
IF objFSO.FileExists(strDesktop & "\������ �� �����.lnk") THEN 
objFSO.DeleteFile(strDesktop & "\������ �� �����.lnk") 
End If 
Set WshShell = NOTHING 
Set objFSO = NOTHING 

Set WshShell = CreateObject ("WScript.Shell")
set objFSO = CreateObject("Scripting.FileSystemObject") 
strDesktop = WshShell.SpecialFolders("Desktop") 
IF objFSO.FileExists(strDesktop & "\� ����� �����.lnk") THEN 
objFSO.DeleteFile(strDesktop & "\� ����� �����.lnk") 
End If 
Set WshShell = NOTHING 
Set objFSO = NOTHING 

Set WshShell = CreateObject ("WScript.Shell")
set objFSO = CreateObject("Scripting.FileSystemObject") 
strDesktop = WshShell.SpecialFolders("Desktop") 
IF objFSO.FileExists(strDesktop & "\���.lnk") THEN 
objFSO.DeleteFile(strDesktop & "\���.lnk") 
End If 
Set WshShell = NOTHING 
Set objFSO = NOTHING 

Set WshShell = CreateObject ("WScript.Shell")
set objFSO = CreateObject("Scripting.FileSystemObject") 
strDesktop = WshShell.SpecialFolders("Desktop") 
IF objFSO.FileExists(strDesktop & "\���������.lnk") THEN 
objFSO.DeleteFile(strDesktop & "\���������.lnk") 
End If 
Set WshShell = NOTHING 
Set objFSO = NOTHING 

Set WshShell = CreateObject ("WScript.Shell")
set objFSO = CreateObject("Scripting.FileSystemObject") 
strDesktop = WshShell.SpecialFolders("Desktop") 
IF objFSO.FileExists(strDesktop & "\���������.lnk") THEN 
objFSO.DeleteFile(strDesktop & "\���������.lnk") 
End If 
Set WshShell = NOTHING 
Set objFSO = NOTHING 


Set WshShell = CreateObject ("WScript.Shell")
set objFSO = CreateObject("Scripting.FileSystemObject") 
strDesktop = WshShell.SpecialFolders("Desktop") 
IF objFSO.FileExists(strDesktop & "\���������.lnk") THEN 
objFSO.DeleteFile(strDesktop & "\���������.lnk") 
End If 
Set WshShell = NOTHING 
Set objFSO = NOTHING 

Set WshShell = CreateObject ("WScript.Shell")
set objFSO = CreateObject("Scripting.FileSystemObject") 
strDesktop = WshShell.SpecialFolders("Desktop") 
IF objFSO.FileExists(strDesktop & "\�����������.lnk") THEN 
objFSO.DeleteFile(strDesktop & "\�����������.lnk") 
End If 
Set WshShell = NOTHING 
Set objFSO = NOTHING 

Set WshShell = CreateObject ("WScript.Shell")
set objFSO = CreateObject("Scripting.FileSystemObject") 
strDesktop = WshShell.SpecialFolders("Desktop") 
IF objFSO.FileExists(strDesktop & "\��������� �� ������.lnk") THEN 
objFSO.DeleteFile(strDesktop & "\��������� �� ������.lnk") 
End If 
Set WshShell = NOTHING 
Set objFSO = NOTHING 

Set WshShell = CreateObject ("WScript.Shell")
set objFSO = CreateObject("Scripting.FileSystemObject") 
strDesktop = WshShell.SpecialFolders("Desktop") 
IF objFSO.FileExists(strDesktop & "\������� �������.lnk") THEN 
objFSO.DeleteFile(strDesktop & "\������� �������.lnk") 
End If 
Set WshShell = NOTHING 
Set objFSO = NOTHING 

Set WshShell = CreateObject ("WScript.Shell")
set objFSO = CreateObject("Scripting.FileSystemObject") 
strDesktop = WshShell.SpecialFolders("Desktop") 
IF objFSO.FileExists(strDesktop & "\����� ������.lnk") THEN 
objFSO.DeleteFile(strDesktop & "\����� ������.lnk") 
End If 
Set WshShell = NOTHING 
Set objFSO = NOTHING 

Set WshShell = CreateObject ("WScript.Shell")
set objFSO = CreateObject("Scripting.FileSystemObject") 
strDesktop = WshShell.SpecialFolders("Desktop") 
IF objFSO.FileExists(strDesktop & "\������� � �������.lnk") THEN 
objFSO.DeleteFile(strDesktop & "\������� � �������.lnk") 
End If 
Set WshShell = NOTHING 
Set objFSO = NOTHING 

Set WshShell = CreateObject ("WScript.Shell")
set objFSO = CreateObject("Scripting.FileSystemObject") 
strDesktop = WshShell.SpecialFolders("Desktop") 
IF objFSO.FileExists(strDesktop & "\Happy 8 marta.lnk") THEN 
objFSO.DeleteFile(strDesktop & "\Happy 8 marta.lnk") 
End If 
Set WshShell = NOTHING 
Set objFSO = NOTHING 

Set WshShell = CreateObject ("WScript.Shell")
set objFSO = CreateObject("Scripting.FileSystemObject") 
strDesktop = WshShell.SpecialFolders("Desktop") 
IF objFSO.FileExists(strDesktop & "\�� ������.lnk") THEN 
objFSO.DeleteFile(strDesktop & "\�� ������.lnk") 
End If 
Set WshShell = NOTHING 
Set objFSO = NOTHING 

set objNetwork = createobject("wscript.network") 
set WShell = WScript.CreateObject("WScript.Shell") 
set fso = createobject("scripting.FileSystemObject") 
pts2="\\192.168.0.100\netlogon\compname"
sname="del_cname.vbs" 
profpath=WShell.SpecialFolders("Desktop")
compname=objNetwork.ComputerName 
LinkName= profpath & "\��� ���������� - " & compname & ".lnk"
   if (fso.fileexists (LinkName) = true) then 
      fso.deletefile (LinkName) 
   end if 

Set WshShell = CreateObject("WScript.Shell")
WshShell.Run "cmd.exe /C ""del /AR %userprofile%\AppData\Local\IconCache.db """, 0, true
WshShell.Run "cmd.exe /C ""del /AS %userprofile%\AppData\Local\IconCache.db """, 0, true
WshShell.Run "cmd.exe /C ""del /AH %userprofile%\AppData\Local\IconCache.db """, 0, true
WshShell.Run "cmd.exe /C ""del /AA %userprofile%\AppData\Local\IconCache.db """, 0, true

Set WshShell = CreateObject ("WScript.Shell")
set objFSO = CreateObject("Scripting.FileSystemObject") 

strDesktop = WshShell.SpecialFolders("Desktop") 


IF objFSO.FileExists(strDesktop & "\������ �� �����.lnk") THEN 

objFSO.DeleteFile(strDesktop & "\������ �� �����.lnk") 

End If 


Set WshShell = NOTHING 
Set objFSO = NOTHING 
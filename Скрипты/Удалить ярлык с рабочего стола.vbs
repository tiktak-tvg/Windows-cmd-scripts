Set WshShell = CreateObject ("WScript.Shell")
set objFSO = CreateObject("Scripting.FileSystemObject") 

strDesktop = WshShell.SpecialFolders("Desktop") 


IF objFSO.FileExists(strDesktop & "\Запись на прием.lnk") THEN 

objFSO.DeleteFile(strDesktop & "\Запись на прием.lnk") 

End If 


Set WshShell = NOTHING 
Set objFSO = NOTHING 
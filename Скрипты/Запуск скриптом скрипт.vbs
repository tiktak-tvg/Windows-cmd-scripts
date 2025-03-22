Set objShell = CreateObject("Shell.Application")
Set FSO = CreateObject("Scripting.FileSystemObject")
strPath = FSO.GetParentFolderName (WScript.ScriptFullName)
If FSO.FileExists(strPath & "\Показывать скрытые файлы и папки.vbs") Then
     objShell.ShellExecute "wscript.exe", _
        Chr(34) & strPath & "\Показывать скрытые файлы и папки.vbs" & Chr(34), "", "runas", 1
Else
     MsgBox "Script file Показывать скрытые файлы и папки.vbs not found"
End If 
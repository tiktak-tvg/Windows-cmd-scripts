Set objShell = CreateObject("Shell.Application")
Set FSO = CreateObject("Scripting.FileSystemObject")
strPath = FSO.GetParentFolderName (WScript.ScriptFullName)
If FSO.FileExists(strPath & "\���������� ������� ����� � �����.vbs") Then
     objShell.ShellExecute "wscript.exe", _
        Chr(34) & strPath & "\���������� ������� ����� � �����.vbs" & Chr(34), "", "runas", 1
Else
     MsgBox "Script file ���������� ������� ����� � �����.vbs not found"
End If 
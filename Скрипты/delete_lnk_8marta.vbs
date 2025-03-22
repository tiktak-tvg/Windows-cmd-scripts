Set Shell = CreateObject("WScript.Shell")
Set FSO = CreateObject("Scripting.FileSystemObject")
DesktopPath = Shell.SpecialFolders("Desktop")
if (FSO.FileExists(DesktopPath & "\Happy 8 marta.lnk")<>false) Then
FSO.DeleteFile DesktopPath & "\Happy 8 marta.lnk"
else
MsgBox("ярлык удалЄн!")
end if

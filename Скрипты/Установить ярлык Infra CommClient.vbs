Dim oShell
Dim oShortCut
set oShell = WScript.CreateObject ("WScript.Shell")
Set objShell = CreateObject("Shell.Application")
objShell.ShellExecute "cmd.exe", " /c cls&&tcmsetup /Q /C INFRA""", , , NORMAL_WINDOW
DesktopPath = oShell.SpecialFolders("AllUsersDesktop")
Set oShortCut = oShell.CreateShortcut(DesktopPath & "\Infra CommClient.lnk")
oShortCut.TargetPath = "C:\Program Files\Infra TeleSystems\InfraClient\BIN\InfraClient.exe"
oShortCut.IconLocation = "C:\Program Files\Infra TeleSystems\InfraClient\BIN\InfraClient.exe"
oShortCut.WorkingDirectory = oShell.SpecialFolders("")
oShortCut.Save()
Set oShell = Nothing
Set oShortCut = Nothing
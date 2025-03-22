Set WshShell = WScript.CreateObject("WScript.Shell")
strDesktop = WshShell.SpecialFolders("AllUsersDesktop")
Set oShellLink = WshShell.CreateShortcut(strDesktop & "\Программы Он клиник.lnk")
oShellLink.TargetPath = "\\192.168.0.100\Froda\soft\ONCLINIC\onclinic.hta"
oShellLink.WindowStyle = 1
oShellLink.IconLocation = "\\192.168.0.100\Froda\soft\ONCLINIC\files\+.ico, 0"
oShellLink.Hotkey = "Ctrl+O"
oShellLink.WorkingDirectory = WshShell.SpecialFolders("AllUsersDesktop")
' Аргументы командной строки:
' Аргументы имеют смысл если файл к которому обращается ярлык - исполняемый, т.е. имеет одно из следующих расширений: .exe; .bat; .com; .cmd; .vbs и т.д.
oShellLink.Arguments = ""
oShellLink.Save
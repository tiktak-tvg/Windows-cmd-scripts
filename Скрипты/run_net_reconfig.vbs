Set WshShell = CreateObject("WScript.Shell")
WshShell.Run "sc stop TermService", 0
Set WshShell = CreateObject("WScript.Shell")
WshShell.Run "ping 127.0.0.1 -n 2", 0, true
Set WshShell = CreateObject("WScript.Shell")
WshShell.Run "sc start TermService", 0
Set WshShell = CreateObject("WScript.Shell")
WshShell.Run "cmd /C ipconfig /flushdns&&ipconfig /release&&ipconfig /renew", 0
Set WshShell = CreateObject("WScript.Shell")
WshShell.Run "netsh interface set interface name=""Подключение по локальной сети"" admin=disabled", 0, true
Set WshShell = CreateObject("WScript.Shell")
WshShell.Run "ping 127.0.0.1 -n 2", 0, true
Set WshShell = CreateObject("WScript.Shell")
WshShell.Run "netsh interface set interface name=""Подключение по локальной сети"" admin=enabled", 0, true
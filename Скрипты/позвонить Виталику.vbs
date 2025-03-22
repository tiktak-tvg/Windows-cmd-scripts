Set WshShell = CreateObject("WScript.Shell")
WshShell.Run "cmd /C C:\Program Files\Infra TeleSystems\InfraClient\BIN\InfraClient.exe", 0, True

Set WSHShell = CreateObject("WScript.Shell")
WSHShell.SendKeys "^(N)"

Set WSHShell = CreateObject("WScript.Shell")
WSHShell.SendKeys "2123"

Set WSHShell = CreateObject("WScript.Shell")
WSHShell.SendKeys "{ENTER}"


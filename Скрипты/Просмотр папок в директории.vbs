Dim oShell
Set oShell = WScript.CreateObject ("WSCript.shell")
oShell.run "cmd /K CD & dir C:\"
oShell.run "cmd /K CD C:\ & Dir"

Set oShell = Nothing
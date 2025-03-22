Dim oShell
Set oShell = WScript.CreateObject ("WSCript.shell")
oShell.run "cmd /K CD C:\ & attrib E:\111 -R -A -S -H /S /D"
Set oShell = Nothing
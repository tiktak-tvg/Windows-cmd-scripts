Option Explicit
Dim WshShell
set WshShell = WScript.CreateObject("WScript.Shell")
WshShell.Run "calc"
WScript.Sleep (20)
WshShell.AppActivate "Калькулятор"
WScript.Sleep (20)
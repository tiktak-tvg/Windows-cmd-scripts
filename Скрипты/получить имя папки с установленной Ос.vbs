Set fso = WScript.CreateObject("Scripting.Filesystemobject")
MsgBox "����� Windows "+fso.GetSpecialFolder(0)
fso.CopyFile "c:\autoexec.bat",fso.GetSpecialFolder(0)+"\"
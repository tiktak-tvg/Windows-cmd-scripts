Set objShell = CreateObject("Shell.Application")
Set objDesktop = objShell.NameSpace(&H0)
desktopDir = objDesktop.Self.Path & "\"
Set objDesktop = Nothing
Set objShell = Nothing
Set objWShell = CreateObject("WScript.Shell")
Set objSC = objWShell.CreateShortcut(desktopDir & "������� � �������.lnk")
objSC.Description = "������� � �������"
objSC.TargetPath = "C:\Program Files\Internet Explorer\iexplore.exe"
objSC.WindowStyle = 1
objSC.WorkingDirectory = "C:\Program Files\Internet Explorer" 
objSC.IconLocation = "C:\Program Files\Internet Explorer\iexplore.exe" & ", 0"
objSC.Arguments = " http://elite-medical.ru/"
objSC.Save()

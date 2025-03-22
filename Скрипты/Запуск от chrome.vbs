Set objShell = CreateObject("Shell.Application")
Set objDesktop = objShell.NameSpace(&H0)
desktopDir = objDesktop.Self.Path & "\"
Set objDesktop = Nothing
Set objShell = Nothing
Set objWShell = CreateObject("WScript.Shell")
Set objSC = objWShell.CreateShortcut(desktopDir & "Клиника в Израиле.lnk")
objSC.Description = "Лечение в Израиле"
objSC.TargetPath = "C:\Program Files (x86)\Google\Chrome\Application\chrome.exe"
objSC.WindowStyle = 1
objSC.WorkingDirectory = "C:\Program Files (x86)\Google\Chrome\Application" 
objSC.IconLocation = "C:\Program Files (x86)\Google\Chrome\Application\chrome.exe" & ", 0"
objSC.Arguments = " http://elite-medical.ru/"
objSC.Save()
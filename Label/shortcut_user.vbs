Set objShell = CreateObject("Shell.Application")
Set objDesktop = objShell.NameSpace(&H0)
desktopDir = objDesktop.Self.Path & "\"
Set objShell = Nothing
Set objWShell = CreateObject("WScript.Shell")
Set objSC = objWShell.CreateShortcut(desktopDir & "Имя компьютера.lnk")
objSC.IconLocation = "F:\oldWin\Documents\Label\name_computer.ico"
objSC.TargetPath = "F:\oldWin\Documents\Label\ComInfo.vbs"
objSC.WindowStyle = 1
objSC.WorkingDirectory = objDesktop.Self.Path & "\"
Set objDesktop = Nothing
objSC.Save
Set objSC = Nothing
Set objWShell = Nothing
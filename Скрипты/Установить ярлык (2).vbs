'************************************************************ 
'* Имя: Shortcut_Create.vbs                                 * 
'* Язык: VBScript                                           * 
'* Назначение: Создание и настройка ярлыка на рабочем столе *
'************************************************************

Set objShell = CreateObject("Shell.Application")
Set objDesktop = objShell.NameSpace(&H0)
desktopDir = objDesktop.Self.Path & "\"
Set objDesktop = Nothing
Set objShell = Nothing
Set objWShell = CreateObject("WScript.Shell")
Set objSC = objWShell.CreateShortcut(desktopDir & "Он Клиник.lnk")
objSC.Description = "Программы"
objSC.IconLocation = "\\192.168.0.100\Froda\soft\ONCLINIC\files\+.ico"
objSC.TargetPath = "\\192.168.0.100\Froda\soft\ONCLINIC\onclinic.hta"
objSC.WindowStyle = 1
objSC.WorkingDirectory = "\\192.168.0.100\Froda\soft\ONCLINIC"
objSC.Save
Set objSC = Nothing
Set objWShell = Nothing
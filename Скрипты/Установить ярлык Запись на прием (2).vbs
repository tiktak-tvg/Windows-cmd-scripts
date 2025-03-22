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
Set objSC = objWShell.CreateShortcut(desktopDir & "Запись на прием.lnk")
objSC.Description = "Программа"
objSC.IconLocation = "\\192.168.0.100\Froda\On_Eib_B\user.ico"
objSC.TargetPath = "\\192.168.0.100\Froda\On_Eib_B\raspis.vbs"
objSC.WindowStyle = 1
objSC.WorkingDirectory = "\\192.168.0.100\Froda\On_Eib_B"
objSC.Save
Set objSC = Nothing
Set objWShell = Nothing
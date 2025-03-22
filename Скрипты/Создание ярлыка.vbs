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
objSC.Description = "План города"
objSC.IconLocation = "shell32.dll, 31"
objSC.TargetPath = "C:\Program Files\7z SFX Tools\7ZSplit.exe"
objSC.WindowStyle = 1
objSC.WorkingDirectory = "C:\Program Files\7z SFX Tools"
objSC.Save
Set objSC = Nothing
Set objWShell = Nothing
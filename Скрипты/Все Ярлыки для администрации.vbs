Dim SavePsw 
Set SavePsw = CreateObject("WScript.Shell")
SavePsw.Run "cmd.exe /c chcp 1251&&net use \\labserver /user:moscow\Администратор_ begin",0,true
SavePsw.Run "cmd.exe /c chcp 1251&&net use \\192.168.0.2 /user:moscow\Администратор_ begin",0,true
SavePsw.Run "cmd.exe /c chcp 1251&&net use \\win2kserv.moscow.onclinic.microsoft.com /user:moscow\Администратор_ begin",0,true
SavePsw.Run "cmd.exe /c chcp 1251&&net use \\192.168.0.100 /user:moscow\Администратор_ begin",0,true
SavePsw.Run "cmd.exe /c chcp 1251&&net use \\bicept.moscow.onclinic.microsoft.com /user:moscow\Администратор_ begin",0,true
SavePsw.Run "cmd.exe /c chcp 1251&&net use \\192.168.0.211 /user:moscow\Администратор_ begin",0,true
SavePsw.Run "cmd.exe /c chcp 1251&&net use /persistent:yes",0,true
Set SavePsw = Nothing

Set objShell = CreateObject("Shell.Application")
Set objDesktop = objShell.NameSpace(&H0)
desktopDir = objDesktop.Self.Path & "\"
Set objDesktop = Nothing
Set objShell = Nothing
Set objWShell = CreateObject("WScript.Shell")
Set objSC = objWShell.CreateShortcut(desktopDir & "Стандарты.lnk")
objSC.Description = "Лечение в Израиле"
objSC.TargetPath = "\\192.168.0.2\Documents\Clinica\Общие документы\Стандарты"
objSC.IconLocation = "\\192.168.0.100\netlogon\label\standart.ico"
objSC.Save
Set objWShell = Nothing

Set objShell = CreateObject("Shell.Application")
Set objDesktop = objShell.NameSpace(&H0)
desktopDir = objDesktop.Self.Path & "\"
Set objDesktop = Nothing
Set objShell = Nothing
Set objWShell = CreateObject("WScript.Shell")
Set objSC = objWShell.CreateShortcut(desktopDir & "Клиника в Израиле.lnk")
objSC.Description = "Лечение в Израиле"
objSC.TargetPath =  "\\192.168.0.100\Rdll\Chrome\Application\chrome.exe"
objSC.IconLocation = "\\192.168.0.100\netlogon\label\Israels.ico"
objSC.Save
objSC.Arguments = " http://elite-medical.ru/"
objSC.Save

Set objShell = CreateObject("Shell.Application")
Set objDesktop = objShell.NameSpace(&H0)
desktopDir = objDesktop.Self.Path & "\"
Set objDesktop = Nothing
Set objShell = Nothing
Set objWShell = CreateObject("WScript.Shell")
Set objSC = objWShell.CreateShortcut(desktopDir & "СТРАХОВЫЕ.lnk")
objSC.Description = "СТРАХОВЫЕ"
objSC.IconLocation = "\\192.168.0.100\netlogon\label\Text.ico"
objSC.TargetPath = "\\192.168.0.2\Documents\Clinica\Кабинеты\этаж 6\кабинет 608\СТРАХОВЫЕ"
objSC.Save
Set objWShell = Nothing



Set objDesktop = Nothing
Set objShell = Nothing
Set objWShell = CreateObject("WScript.Shell")
Set objSC = objWShell.CreateShortcut(desktopDir & "Новости клиники.lnk")
objSC.Description = "Новости клиники"
objSC.TargetPath = "\\192.168.0.2\Documents\Clinica\Новости клиники"
objSC.IconLocation = "\\192.168.0.100\netlogon\label\Новости клиники.ico"
objSC.Save
Set objWShell = Nothing


Set objShell = CreateObject("Shell.Application")
Set objDesktop = objShell.NameSpace(&H0)
desktopDir = objDesktop.Self.Path & "\"
Set objDesktop = Nothing
Set objShell = Nothing
Set objWShell = CreateObject("WScript.Shell")
Set objSC = objWShell.CreateShortcut(desktopDir & "Лечение в Израиле.lnk")
objSC.Description = "Лечение в Израиле"
objSC.TargetPath = "\\192.168.0.2\Documents\Clinica\Общие документы\Лечение в Израиле"
objSC.IconLocation = "\\192.168.0.100\netlogon\label\Israel.ico"
objSC.Save
Set objWShell = Nothing

Set objShell = CreateObject("Shell.Application")
Set objDesktop = objShell.NameSpace(&H0)
desktopDir = objDesktop.Self.Path & "\"
Set objDesktop = Nothing
Set objShell = Nothing
Set objWShell = CreateObject("WScript.Shell")
Set objSC = objWShell.CreateShortcut(desktopDir & "Запись на прием.lnk")
objSC.Description = "Программа"
objSC.IconLocation = "\\192.168.0.100\netlogon\label\user.ico"
objSC.TargetPath = "\\192.168.0.211\Froda\On_Eib_C\raspis.vbs"
objSC.WindowStyle = 1
objSC.WorkingDirectory = "\\192.168.0.211\Froda\On_Eib_C"
objSC.Save
Set objSC = Nothing
Set objWShell = Nothing

Set objShell = CreateObject("Shell.Application")
Set objDesktop = objShell.NameSpace(&H0)
desktopDir = objDesktop.Self.Path & "\"
Set objDesktop = Nothing
Set objShell = Nothing
Set objWShell = CreateObject("WScript.Shell")
Set objSC = objWShell.CreateShortcut(desktopDir & "ЭИБ.lnk")
objSC.Description = "Программа"
objSC.IconLocation = "\\192.168.0.100\netlogon\label\TREESURG.ICO"
objSC.TargetPath = "\\192.168.0.211\Froda\On_Eib_C\eib.vbs"
objSC.WindowStyle = 1
objSC.WorkingDirectory = "\\192.168.0.211\Froda\On_Eib_C"
objSC.Save
Set objSC = Nothing
Set objWShell = Nothing

Set objShell = CreateObject("Shell.Application")
Set objDesktop = objShell.NameSpace(&H0)
desktopDir = objDesktop.Self.Path & "\"
Set objDesktop = Nothing
Set objShell = Nothing
Set objWShell = CreateObject("WScript.Shell")
Set objSC = objWShell.CreateShortcut(desktopDir & "Лаборатория.lnk")
objSC.Description = "Программа"
objSC.IconLocation = "\\192.168.0.100\netlogon\label\blood.ico"
objSC.TargetPath = "\\192.168.0.211\Froda\On_Eib_C\pusk.vbs"
objSC.WindowStyle = 1
objSC.WorkingDirectory = "\\192.168.0.211\Froda\On_Eib_C"
objSC.Save
Set objSC = Nothing
Set objWShell = Nothing


Set objShell = CreateObject("Shell.Application")
Set objDesktop = objShell.NameSpace(&H0)
desktopDir = objDesktop.Self.Path & "\"
Set objDesktop = Nothing
Set objShell = Nothing
Set objWShell = CreateObject("WScript.Shell")
Set objSC = objWShell.CreateShortcut(desktopDir & "Бланк заявки.lnk")
objSC.Description = "Бланк заявки"
objSC.IconLocation = "\\192.168.0.100\netlogon\label\Edit Text.ico"
objSC.TargetPath = "\\192.168.0.2\For admin\Pub\Бланк заявки"
objSC.Save
Set objWShell = Nothing

Set objShell = CreateObject("Shell.Application")
Set objDesktop = objShell.NameSpace(&H0)
desktopDir = objDesktop.Self.Path & "\"
Set objDesktop = Nothing
Set objShell = Nothing
Set objWShell = CreateObject("WScript.Shell")
Set objSC = objWShell.CreateShortcut(desktopDir & "Программы Он Клиник.lnk")
objSC.Description = "Программа"
objSC.IconLocation = "\\192.168.0.100\netlogon\label\onclinic.ico"
objSC.TargetPath = "\\192.168.0.211\Froda\On_Eib_C\Программы\Цветной Бульвар.hta"
objSC.WindowStyle = 1
objSC.WorkingDirectory = "\\192.168.0.211\Froda\On_Eib_C\Программы"
objSC.Save
Set objSC = Nothing
Set objWShell = Nothing

Set objShell = CreateObject("Shell.Application")
Set objDesktop = objShell.NameSpace(&H0)
desktopDir = objDesktop.Self.Path & "\"
Set objDesktop = Nothing
Set objShell = Nothing
Set objWShell = CreateObject("WScript.Shell")
Set objSC = objWShell.CreateShortcut(desktopDir & "Список телефонов.lnk")
objSC.Description = "Список телефонов"
objSC.IconLocation = "\\192.168.0.100\netlogon\label\numbers.ico"
objSC.TargetPath = "\\192.168.0.211\Froda\On_Eib_C\Программы\spisok.hta"
objSC.WindowStyle = 1
objSC.WorkingDirectory = "\\192.168.0.211\Froda\On_Eib_C\Программы"
objSC.Save
Set objSC = Nothing
Set objWShell = Nothing


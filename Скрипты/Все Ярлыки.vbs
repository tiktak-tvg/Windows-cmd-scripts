' доступ ко всем серверам с проверкой на существование cmdkey.exe - windows7
Dim Filesys, SavePsw, SavePsw1 
Set Filesys = CreateObject("Scripting.FileSystemObject") 
If Filesys.FileExists("%windir%\System32\cmdkey.exe") Then
	Set SavePsw = CreateObject("WScript.Shell")
		SavePsw.Run "cmd.exe /c chcp 1251&&cmdkey.exe /add:192.168.0.210 /user:moscow\Администратор_ /pass:begin",0,true
		SavePsw.Run "cmd.exe /c chcp 1251&&cmdkey.exe /add:bicepz /user:moscow\Администратор_ /pass:begin",0,true
		SavePsw.Run "cmd.exe /c chcp 1251&&cmdkey.exe /add:192.168.0.204 /user:moscow\Администратор_ /pass:begin",0,true
		SavePsw.Run "cmd.exe /c chcp 1251&&cmdkey.exe /add:bak-server /user:moscow\Администратор_ /pass:begin",0,true
		SavePsw.Run "cmd.exe /c chcp 1251&&cmdkey.exe /add:192.168.0.100 /user:moscow\Администратор_ /pass:begin",0,true
		SavePsw.Run "cmd.exe /c chcp 1251&&cmdkey.exe /add:win2kserv1 /user:moscow\Администратор_ /pass:begin",0,true
		SavePsw.Run "cmd.exe /c chcp 1251&&cmdkey.exe /add:192.168.0.205 /user:moscow\Администратор_ /pass:begin",0,true
		SavePsw.Run "cmd.exe /c chcp 1251&&cmdkey.exe /add:lab-server /user:moscow\Администратор_ /pass:begin",0,true
	Set SavePsw = Nothing
Else
	Set SavePsw1 = CreateObject("WScript.Shell")
		SavePsw1.Run "cmd.exe /c chcp 1251&&net use \\bak-server /user:moscow\Администратор_ begin",0,true
		SavePsw1.Run "cmd.exe /c chcp 1251&&net use \\192.168.0.204 /user:moscow\Администратор_ begin",0,true
		SavePsw1.Run "cmd.exe /c chcp 1251&&net use \\lab-server /user:moscow\Администратор_ begin",0,true
		SavePsw1.Run "cmd.exe /c chcp 1251&&net use \\192.168.0.205 /user:moscow\Администратор_ begin",0,true
		SavePsw1.Run "cmd.exe /c chcp 1251&&net use \\bicepz /user:moscow\Администратор_ begin",0,true
		SavePsw1.Run "cmd.exe /c chcp 1251&&net use \\192.168.0.210 /user:moscow\Администратор_ begin",0,true
		SavePsw1.Run "cmd.exe /c chcp 1251&&net use \\win2kserv1 /user:moscow\Администратор_ begin",0,true
		SavePsw1.Run "cmd.exe /c chcp 1251&&net use \\192.168.0.100 /user:moscow\Администратор_ begin",0,true
		SavePsw1.Run "cmd.exe /c net use /persistent:yes",0,true
	Set SavePsw1 = Nothing
End If

' ярлыки

Set objShell = CreateObject("Shell.Application")
Set objDesktop = objShell.NameSpace(&H0)
desktopDir = objDesktop.Self.Path & "\"
Set objDesktop = Nothing
Set objShell = Nothing
Set objWShell = CreateObject("WScript.Shell")
Set objSC = objWShell.CreateShortcut(desktopDir & "Запись на прием.lnk")
objSC.Description = "Программа"
objSC.IconLocation = "\\192.168.0.100\Froda\On_Eib_T\user.ico"
objSC.TargetPath = "\\192.168.0.100\Froda\On_Eib_T\raspis.vbs"
objSC.WindowStyle = 1
objSC.WorkingDirectory = "\\192.168.0.100\Froda\On_Eib_T"
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
objSC.IconLocation = "\\192.168.0.100\Froda\On_Eib_T\TREESURG.ICO"
objSC.TargetPath = "\\192.168.0.100\Froda\On_Eib_T\eib.vbs"
objSC.WindowStyle = 1
objSC.WorkingDirectory = "\\192.168.0.100\Froda\On_Eib_T"
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
objSC.IconLocation = "\\192.168.0.100\Froda\On_Eib_T\blood.ico"
objSC.TargetPath = "\\192.168.0.100\Froda\On_Eib_T\pusk.vbs"
objSC.WindowStyle = 1
objSC.WorkingDirectory = "\\192.168.0.100\Froda\On_Eib_T"
objSC.Save
Set objSC = Nothing
Set objWShell = Nothing
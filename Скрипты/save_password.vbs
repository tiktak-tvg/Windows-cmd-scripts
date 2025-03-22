Dim SavePsw 
Set SavePsw = CreateObject("WScript.Shell")
SavePsw.Run "cmd.exe /c chcp 1251&&net use \\dc1.infratel.onclnic.ru /user:infratel\infra infra",0,true
SavePsw.Run "cmd.exe /c chcp 1251&&net use \\192.168.101.101 /user:infratel\infra infra",0,true
SavePsw.Run "cmd.exe /c chcp 1251&&net use /persistent:yes",0,true
Set SavePsw = Nothing
Set SavePsw = Nothing
	Set obNetwork = CreateObject("WScript.Network")
	Set oDrives = obNetwork.EnumNetworkDrives
		mydrv = "Y:"
		mapped = false
		myshare = "\\192.168.101.101\MailKid"
		For i = 0 to oDrives.Count - 1 Step 2
			  If oDrives.Item(i)=mydrv Then mapped = true
		Next
If Not mapped Then obNetwork.MapNetworkDrive mydrv, myshare

Set objShell = CreateObject("Shell.Application")
Set objDesktop = objShell.NameSpace(&H0)
desktopDir = objDesktop.Self.Path & "\"
Set objDesktop = Nothing
Set objShell = Nothing
Set objWShell = CreateObject("WScript.Shell")
Set objSC = objWShell.CreateShortcut(desktopDir & "EMailDt.lnk")
objSC.Description = "Автоответчик"
objSC.TargetPath = "Y:\"
objSC.IconLocation = "\\192.168.0.100\netlogon\label\Email.ico"
objSC.Save
Set objWShell = Nothing


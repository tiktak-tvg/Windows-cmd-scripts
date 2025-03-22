'Этим скриптом можно удалить любую установленную через Windows Installer программу. В этом примере её имя LeftSoft Program.
strComputer = "."
Set objWMIService = GetObject("winmgmts:" _
   & "{impersonationLevel=impersonate}!\\" & strComputer & "\root\cimv2")

Set colSoftware = objWMIService.ExecQuery _
   ("Select * from Win32_Product Where Name = 'LeftSoft Program'")

For Each objSoftware in colSoftware
   objSoftware.Uninstall()
Next
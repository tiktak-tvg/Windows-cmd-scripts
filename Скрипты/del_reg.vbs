const HKEY_LOCAL_MACHINE = &H80000002
strComputer = "."
Set StdOut = WScript.StdOut
 
Set oReg=GetObject("winmgmts:{impersonationLevel=impersonate}!\\" &_ 
strComputer & "\root\default:StdRegProv")
 
strKeyPath = "HKEY_CLASSES_ROOT\Directory\Background\shell\Имя и IP адрес компьютера"
oReg.CreateKey HKEY_LOCAL_MACHINE,strKeyPath
' Сохранить как EnableDisableNIC.vbs

strComputer = "."
Set objWMIService = GetObject("winmgmts:" & "{impersonationLevel=impersonate}!\\" & strComputer & "\root\cimv2")

Set objNetworkSettings = objWMIService.Get("Win32_NetworkAdapter ")
If objNetworkSettings.Disable()=0 then
'If objNetworkSettings.Enable()=0 then
    MsgBox "Succes"
else
    MsgBox "Error"
end if
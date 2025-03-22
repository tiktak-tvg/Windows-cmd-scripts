'¬ключает службу DHCP-клиента на компьютере. компьютер будет использовать DHCP дл€ получени€ IP-адреса, а не использовать статический IP-адрес.
strComputer = "." 
Set objWMIService = GetObject("winmgmts:" _ 
    & "{impersonationLevel=impersonate}!\\" & strComputer & "\root\cimv2") 
 
Set colNetAdapters = objWMIService.ExecQuery _ 
    ("Select * from Win32_NetworkAdapterConfiguration where IPEnabled=TRUE") 
  
For Each objNetAdapter In colNetAdapters 
    errEnable = objNetAdapter.EnableDHCP() 
Next 

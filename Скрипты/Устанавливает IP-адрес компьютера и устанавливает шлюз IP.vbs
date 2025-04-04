strComputer = "." 
Set objWMIService = GetObject("winmgmts:" _ 
    & "{impersonationLevel=impersonate}!\\" & strComputer & "\root\cimv2") 
 
Set colNetAdapters = objWMIService.ExecQuery _ 
    ("Select * from Win32_NetworkAdapterConfiguration where IPEnabled=TRUE") 
 
strIPAddress = Array("192.168.1.141") 
strSubnetMask = Array("255.255.255.0") 
strGateway = Array("192.168.1.100") 
strGatewayMetric = Array(1) 
  
For Each objNetAdapter in colNetAdapters 
    errEnable = objNetAdapter.EnableStatic(strIPAddress, strSubnetMask) 
    errGateways = objNetAdapter.SetGateways(strGateway, strGatewaymetric) 
    If errEnable = 0 Then 
        WScript.Echo "The IP address has been changed." 
    Else 
        WScript.Echo "The IP address could not be changed." 
    End If 
Next 

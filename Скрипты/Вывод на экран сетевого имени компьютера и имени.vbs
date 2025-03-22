On Error Resume Next
strComputer = "localhost"
Set objWMIService = GetObject("winmgmts:" & "{impersonationLevel=impersonate}!\\" & strComputer & "\root\cimv2")
regValueDataMetric = "35"

'Set colItems = objWMIService.ExecQuery ("Select * From Win32_NetworkAdapter ")
'Where NetConnectionID = 'Wireless Network Connection'

'For Each objItem in colItems
'strMACAddress = objItem.MACAddress
'Wscript.Echo "MACAddress: " & strMACAddress
'Next

Set colNetCard = objWMIService.ExecQuery ("Select * From Win32_NetworkAdapterConfiguration where IPEnabled=TRUE")

For Each objNetCard in colNetCard
'If objNetCard.MACAddress = strMACAddress Then
For Each strIPAddress in objNetCard.IPAddress
Wscript.Echo "Description: " & objNetCard.Description
Wscript.Echo "IP Address: " & strIPAddress
Wscript.Echo "IPConnectionMetric: " & objNetCard.IPConnectionMetric
objNetCard.SetIPConnectionMetric(regValueDataMetric)
Next
'End If
Next
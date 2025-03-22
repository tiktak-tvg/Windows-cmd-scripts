Set WMI = GetObject("winmgmts:\\.\root\cimv2")
Set oDGs = WMI.ExecQuery("Select * from Win32_NetworkAdapterConfiguration where IPEnabled=TRUE")
For Each oDG In oDGs
If Not IsNull(oDG.DefaultIPGateway) Then
DefaultGateway=oDG.DefaultIPGateway(0)
End If
Next
'*************************************************************************
'* Имя: GetIPaddressOfParticularConnection.vbs                           * 
'* Язык: VBScript                                                        * 
'* Назначение: Получить IP-адрес для определённого сетевого соединения   *
'*************************************************************************

strComputer  =  "."
strNetworkConnection = "'Подключение по локальной сети'" ' <- редактировать под нужное 
                                                 '    имя сетевого конекта

Set objWMIService = GetObject( _
    "winmgmts:\\" & strComputer & "\root\cimv2")
Set colNics = objWMIService.ExecQuery _
    ("Select * From Win32_NetworkAdapter " _
        & "Where NetConnectionID = "  &   strNetworkConnection)
 
For Each objNic in colNics
    Set colNicConfigs = objWMIService.ExecQuery _
      ("ASSOCIATORS OF " _
          & "{Win32_NetworkAdapter.DeviceID='" & _
      objNic.DeviceID & "'}" & _
      " WHERE AssocClass=Win32_NetworkAdapterSetting")
    For Each objNicConfig In colNicConfigs
        For Each strIPAddress in objNicConfig.IPAddress
            Wscript.Echo "IP Address: " &  strIPAddress
        Next
    Next
Next
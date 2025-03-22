'*************************************************************************
'* Имя: SetAllNetworkAdapters2DHCP.vbs                                   * 
'* Язык: VBScript                                                        * 
'* Назначение: Установить на всех сетевых адаптерах сетевую конфигурацию *
'*             через DHCP                                                *
'*************************************************************************

strComputer = "."

Set objWMIService = GetObject(_
    "winmgmts:\\" & strComputer & "\root\cimv2")
Set colNetAdapters = objWMIService.ExecQuery _
    ("Select * from Win32_NetworkAdapterConfiguration " _
        & "where IPEnabled=TRUE")
 
For Each objNetAdapter In colNetAdapters
    errEnable = objNetAdapter.EnableDHCP()
Next
on Error Resume Next 

function ipLocal()
xz=""
dim strComputer,objWMIService,IPConfigSet,IPConfig,i,ip
Set WSHNetwork = CreateObject("WScript.Network")
ComputerName=WSHNetwork.ComputerName
UserName = WSHNetwork.UserName

			strComputer = "."
			Set objWMIService = GetObject("winmgmts:\\" & strComputer & "\root\cimv2")
			'Set IPConfigSet = objWMIService.ExecQuery("Select IPAddress from Win32_NetworkAdapterConfiguration where IPEnabled=TRUE")
Set IPConfigSet = objWMIService.ExecQuery("select * from win32_networkadapterconfiguration WHERE IPEnabled='TRUE' " _
   & "AND ServiceName<>'AsyncMac' " _ 
   & "AND ServiceName<>'VMnetx' " _
   & "AND ServiceName<>'VMnetadapter' " _
   & "AND ServiceName<>'Rasl2tp' " _
   & "AND ServiceName<>'msloop' " _ 
   & "AND ServiceName<>'PptpMiniport' " _ 
   & "AND ServiceName<>'Raspti' " _
   & "AND ServiceName<>'NDISWan' " _
   & "AND ServiceName<>'NdisWan4' " _ 
   & "AND ServiceName<>'RasPppoe' " _
   & "AND ServiceName<>'NdisIP' " _ 
   & "AND ServiceName<>'' " _ 
   & "AND Description<>'PPP Adapter.'",,48)
			'определение своего ip
			'For Each IPConfig in IPConfigSet
				'If Not IsNull(IPConfig.IPAddress) Then
					'For i=LBound(IPConfig.IPAddress) to UBound(IPConfig.IPAddress)
					'if len(IPConfig.IPAddress(i))<16 then
					'if IPConfig.IPAddress(i)<>xz then
					'xz=IPConfig.IPAddress(i)
					'ip = ip & chr(13) & IPConfig.IPAddress(i)
					'end if
					'end if
					'Next
				'End If
			'Next
For Each IPConfig in IPConfigSet
   count_all = count_all + 1
 
   if IPConfig.IPAddress(0) <> "0.0.0.0" then
     count = count + 1
     if count = 1 then
        IP = IPConfig.IPAddress(0)
        Subnet = IPConfig.IPSubnet(0)
     end if
   end if
Next
			ipLocal = "Ваш IP адресс: " & IP & chr(13) &  chr(13) & "Имя Компьютера: " & ComputerName & chr(13) & "Имя пользователя: " & UserName & chr(13) & "-----------------------"& chr(13) &"   Телефоны:"& chr(13) &"Тех.поддержка: +7 (906)750-77-59"
End function

msgbox  ipLocal
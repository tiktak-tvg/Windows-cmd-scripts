strGroupName = "New_group"
strCompName = "WS_name"
strDomainName = "Domain_name"
Set objComputer = GetObject("WinNT://" & strDomainName & "/" & strCompName)
Set objGroup = objComputer.Create("group", strGroupName)
objGroup.SetInfo
Set objGroup = Nothing
Set objComputer = Nothing
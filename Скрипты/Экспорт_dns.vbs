Dim IPAdd 
Dim Computer 
strComputer = "win2ksrv0" 

Set objFSO = CreateObject("Scripting.FileSystemObject") 
Set objTextFile = objFSO.CreateTextFile("d:\dns.txt") 


Set objWMIService = GetObject("winmgmts:\\" & strComputer & "\root\MicrosoftDNS") 
Set colItems = objWMIService.ExecQuery( _ 
    "SELECT * FROM MicrosoftDNS_PTRType",,48) 
For Each objItem in colItems 

    CompNameArray = Split(objItem.RecordData ,  ".") 
    For i = LBound(CompNameArray) to UBound(CompNameArray) 
    Computer = CompNameArray(0) 
    Next 

    objTextFile.WriteLine  "\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\" 
    objTextFile.WriteLine  "DNS Info for - " & UCase(Computer) 
    objTextFile.WriteLine  "////////////////////////////////////" & vbCrLf 
    objTextFile.WriteLine  "DNSServerName: " & objItem.DnsServerName 

    IPAddArray = Split (objItem.OwnerName , ".") 
     
    For i = LBound(IPAddArray) to UBound(IPAddArray) 
    IPAdd = IPAddArray(3) & "." & IPAddArray(2) & "." & IPAddArray(1) & "." & IPAddArray(0) 
    Next 
     
    objTextFile.WriteLine  "IP Address: " & IPAdd 
    objTextFile.WriteLine  "FullyQualifiedDomainName: " & objItem.RecordData 
    objTextFile.WriteLine  "Timestamp: " & objItem.Timestamp & vbCrLf 

Next 


objTextFile.Close 
Set objFSO = Nothing 
Set objFile = Nothing 
Set objWMIService = Nothing 
Set colItems = Nothing 
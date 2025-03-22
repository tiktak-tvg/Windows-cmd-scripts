strComputer = "."
Set objWMIService = GetObject("winmgmts:" _
    & "{impersonationLevel=impersonate,(Security)}!\\" & _
        strComputer & "\root\cimv2")

Set colLogFiles = objWMIService.ExecQuery _
    ("Select * from Win32_NTEventLogFile where LogFileName='Security'")

For Each objLogFile in colLogFiles
    Wscript.Echo objLogFile.NumberOfRecords
    Wscript.Echo "Максимальный размер: " &  objLogfile.MaxFileSize 
Next
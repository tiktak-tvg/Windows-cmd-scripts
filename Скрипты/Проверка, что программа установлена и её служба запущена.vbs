strComputer = "."
 Set objWMIService = GetObject("winmgmts:" _
     & "{impersonationLevel=impersonate}!\\" & strComputer & "\root\cimv2")
  Set colServices = objWMIService.ExecQuery _
     ("Select * from Win32_Service Where Name = 'MSSQLServer'")
  If colServices.Count > 0 Then
     For Each objService in colServices
         Wscript.Echo "Служба SQL Server " & objService.State & "."
     Next
 Else
     Wscript.Echo "SQL Server не установлен на этом компьютере."
 End If 
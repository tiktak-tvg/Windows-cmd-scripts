strComputer = "." 
Set objWMIService = GetObject("winmgmts:" _ 
 & "{impersonationLevel=impersonate}!\\" & strComputer & "\root\cimv2") 
Set colItems = objWMIService.ExecQuery("Select * from Win32_LocalTime") 
For Each objItem in colItems 
    Wscript.Echo "Year: " & objItem.Year 
    Wscript.Echo "Month: " & objItem.Month 
    Wscript.Echo "Day: " & objItem.Day 
    Wscript.Echo "Hour: " & objItem.Hour 
    Wscript.Echo "Minute: " & objItem.Minute 
    Wscript.Echo "Second: " & objItem.Second 
    Wscript.Echo "Quarter: " & objItem.Quarter 
    Wscript.Echo "Week in the Month: " & objItem.WeekInMonth 
    Wscript.Echo "Day of the Week: " & objItem.DayOfWeek 
Next 

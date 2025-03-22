Set objShell = CreateObject("Wscript.Shell")
Set objNet = CreateObject("Wscript.Network")

intNumberID = 528
intEvent = 1
intRecordNum = 1

strComputer = "."
strFileName = "Logintimes.txt"
strFolder = "C:\"
strPath = strFolder & strFileName

endTime = Now()

Set objFso = CreateObject("Scripting.FileSystemObject")

If objFSO.FolderExists(strFolder) Then
    Set objFolder = objFSO.GetFolder(strFolder)
Else
   Set objFolder = objFSO.CreateFolder(strFolder)
   Wscript.Echo "Folder created " & strFolder
End If
Set strFile = objFso.OpenTextFile(strPath, 8, True)

Set objWMIService = GetObject("winmgmts:" _
    & "{impersonationLevel=impersonate,(Security)}!\" & _
        strComputer & "\root\cimv2")
Set colLogFiles = objWMIService.ExecQuery _
    ("SELECT * FROM Win32_NTLogEvent WHERE LogFile=`Security`")

intEvent = 1
hightime = 0

For Each objEvent in colLogFiles

If objEvent.EventCode = intNumberID AND (NOT objEvent.User = "NT AUTHORITY\NETWORK SERVICE") Then

startTime = WMIDateStringToDate(objEvent.TimeWritten)

LoginTime = DateDiff("s", startTime, endTime)

strLogin =  "User: " & objEvent.User & vbCRLF & "Start Time: " & startTime & vbCRLF & "End time: " & endTime & vbCRLF & vbCRLF & "Login time: " & LoginTime & " seconds"

strFile.WriteLine (strLogin)
strFile.WriteLine ()

Wscript.Echo strLogin

objShell.LogEvent 4,"User Logon Finished: " & objNet.Username & vbCRLf & strLogin

Wscript.Quit
End if

Next


Function WMIDateStringToDate(dtmInstallDate)
    WMIDateStringToDate = CDate(Mid(dtmInstallDate, 7, 2) & "/" & _
         Mid(dtmInstallDate, 5, 2) & "/" & Left(dtmInstallDate, 4) _
             & " " & Mid (dtmInstallDate, 9, 2) & ":" & _
                 Mid(dtmInstallDate, 11, 2) & ":" & Mid(dtmInstallDate, _
                     13, 2))
End Function

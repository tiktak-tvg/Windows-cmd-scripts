' VBScript Source File -- Created with SAPIEN Technologies PrimalScript Pro 4.0 
' NAME: Password Change - by file listing the systems 
' AUTHOR: Steven W. Hamlin, hamlinsw@msn.com 
' DATE  : 9/23/05 
' Version: 1.01 2/28/06 - updated comments 
' Assumptions: 
'    1)  The file systems.txt exists in the same folder as the script 
'    2)  The systems.txt file is of type ASCII text. 
'    3)  The file systems.txt contains the NETBIOS names or IP addresses of  
'        each system to be updated, presented one system to a line. 
'    4)  There is sufficient room on the drive containing the script To 
'        allow the creation of the log file. 
 
Option Explicit 
 
Dim objIE, objIE2, strDebug, strMessage, arrComputers, strComputer 
Dim objWMIService, objItem, colItems, strLogging, intCount, objFSO 
Dim objLogFile, errReturnCode1, objPing, objStatus, objFSO2 
Dim objReadFile, strPingStatus, strReturnCode1, strReturnCode2, objUser 
Dim strUser, strPwd 
strLogging = "On" 
intCount = 1 
On Error Resume Next 
Const wbemFlagReturnImmediately = &h10 
Const wbemFlagForwardOnly = &h20 
Const ForReading = 1 
strUser = InputBox("Enter the username:") 
strPwd = InputBox("Enter the new password:") 
Record "Computer, Action" 
WScript.Echo strUser & ", " & strPwd 
Set objFSO2 = CreateObject("Scripting.FileSystemObject") 
Set objReadFile = objFSO2.OpenTextFile("systems.txt", ForReading) 
Do Until objReadFile.AtEndOfStream 
    strComputer = objReadFile.ReadLine 
    Set objPing = GetObject("winmgmts:{impersonationLevel=impersonate}")._ 
        ExecQuery("select * from Win32_PingStatus where address = '"_ 
        & strComputer & "'") 
    For Each objStatus in objPing 
        Select Case objStatus.StatusCode 
            Case 0 strPingStatus = "Success"  
            Case 11001 strPingStatus = "Buffer Too Small " 
            Case 11002 strPingStatus = "Destination Net Unreachable " 
            Case 11003 strPingStatus = "Destination Host Unreachable" 
            Case 11004 strPingStatus = "Destination Protocol Unreachable" 
            Case 11005 strPingStatus = "Destination Port Unreachable" 
            Case 11006 strPingStatus = "No Resources" 
            Case 11007 strPingStatus = "Bad Option" 
            Case 11008 strPingStatus = "Hardware Error" 
            Case 11009 strPingStatus = "Packet Too Big" 
            Case 11010 strPingStatus = "Request Timed Out" 
            Case 11011 strPingStatus = "Bad Request" 
            Case 11012 strPingStatus = "Bad Route" 
            Case 11013 strPingStatus = "TimeToLive Expired Transit" 
            Case 11014 strPingStatus = "TimeToLive Expired Reassembly" 
            Case 11015 strPingStatus = "Parameter Problem" 
            Case 11016 strPingStatus = "Source Quench" 
            Case 11017 strPingStatus = "Option Too Big"  
            Case 11018 strPingStatus = "Bad Destination"  
            Case 11032 strPingStatus = "Negotiating IPSEC"  
            Case 11050 strPingStatus = "General Failure" 
        End Select 
        WScript.Echo strComputer & " Ping Status: " & strPingStatus 
        If IsNull(objStatus.StatusCode) Or objStatus.StatusCode<>0 Then  
            Record strComputer & ",Not Avail." 
        Else 
            Set objUser = GetObject("WinNT://" & strComputer & _ 
              "/" & strUser & ", user") 
            'reset the password 
            objUser.SetPassword strPwd 
            'save the change 
            objUser.SetInfo 
            Record strComputer & ",Password Set" 
        End If 
    Next 
Loop 
objReadFile.Close 
Function WMIDateStringToDate(dtmDate) 
WScript.Echo dtm:  
    WMIDateStringToDate = CDate(Mid(dtmDate, 5, 2) & "/" & _ 
    Mid(dtmDate, 7, 2) & "/" & Left(dtmDate, 4) _ 
    & " " & Mid (dtmDate, 9, 2) & ":" & Mid(dtmDate, 11, 2) & ":" & Mid(dtmDate,13, 2)) 
End Function 
Sub Record (strMessage) 
    If strLogging = "On" And intCount = 1 Then 
        Set objFSO = WScript.CreateObject("Scripting.FileSystemObject") 
        Set objLogFile = objFSO.CreateTextFile("log.csv", True) 
        intCount = intCount + 1 
    End If 
    If strLogging = "On" Then 
        objLogFile.WriteLine strMessage 
    End If 
End Sub 
objLogFile.Close 

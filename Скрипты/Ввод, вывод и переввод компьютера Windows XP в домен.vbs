'Script name: NetDomOU.vbs
'Author: Geert Van Camp
'Description: joins,unjoins or removes a Windows XP computer to or from the domain.
'        An organisational unit can be specified when joining a domain!
'Requirements:
'    - only works on Windows XP (and higher)
'    - run the script on the computer that needs to be joined/unjoined/removed
'    - specify a domain user account that has permissions to join/unjoin computers to/from the domain
'    - run with CScript.exe to view output

On Error Resume Next
Const JOIN_DOMAIN = 1
Const ACCT_CREATE = 2
Const ACCT_DOMAIN_JOIN_IF_JOINED = 32

varExitErrorLevel = 0

If WScript.Arguments.Length > 0 Then
    Select Case LCase(WScript.Arguments(0))
        Case "join" strCommand = "Join"
        Case "unjoin" strCommand = "Unjoin"
        Case "remove" strCommand = "Remove"
        Case Else subUsage
    End Select
    For varIndex = 1 To (WScript.Arguments.Length - 1)
        arrArgument = Split(WScript.Arguments(varIndex), ":", -1, vbTextCompare)
        strArgument = arrArgument(0)
        If Ubound(arrArgument) = 0 Then
            Select Case LCase(strArgument)
                Case "/reboot" flgReboot = True
                Case Else subUsage
            End Select
        Else
            strArgumentValue = arrArgument(1)
            Select Case LCase(strArgument)
                Case "/domain" strDomain = strArgumentValue
                Case "/ou" strOU = strArgumentValue
                Case "/user" strUser = strArgumentValue
                Case "/password" strPassword = strArgumentValue
                Case Else subUsage
            End Select
        End If
    Next
Else
    subUsage
End If

Set objNetwork = CreateObject("WScript.Network")
strHostName = objNetwork.ComputerName
Set objNetwork = Nothing

Set objWMIComputer = GetObject("winmgmts:{impersonationLevel=Impersonate}!\\" & strHostName & "\root\cimv2:Win32_ComputerSystem.Name='" & strHostName & "'")
If Err = 0 Then
    Select Case strCommand
        Case "Join"
            subDisplay "Joining computer to domain." & vbCrLf & "Hostname: " & strHostName & vbCrLf & "Domain: " & strDomain & vbCrLf & "OU: " & strOU & vbCrLf & "Username: " & strUser
            varWMIJoinReturnValue = objWMIComputer.JoinDomainOrWorkGroup(strDomain, strPassword, strUser, strOU, JOIN_DOMAIN + ACCT_CREATE)
            If Err = 0 Then
                If varWMIJoinReturnValue = 2224 Then
                    subDisplay "The computer account already exists."
                    If Not strOU = "" Then subDisplay "The computer account will stay in it's current OU."
                    varWMIJoinReturnValue = objWMIComputer.JoinDomainOrWorkGroup(strDomain, strPassword, strUser, strOU, JOIN_DOMAIN)
                    If Not varWMIJoinReturnValue = 0 Then subDisplay fncErrorMessage(varWMIJoinReturnValue, "", True)
                Else
                    If Not varWMIJoinReturnValue = 0 Then subDisplay fncErrorMessage(varWMIJoinReturnValue, "", True)
                End If
            Else
                subDisplay fncErrorMessage(Hex(Err.Number), Err.Description, True)
            End If
        Case "Unjoin", "Remove"
            subDisplay "Unjoining the computer from the domain."
            varWMIJoinReturnValue = objWMIComputer.UnJoinDomainOrWorkGroup( , , 0)
            If Err = 0 Then
                If Not varWMIJoinReturnValue = 0 Then subDisplay fncErrorMessage(varWMIJoinReturnValue, "", True)
            Else
                subDisplay fncErrorMessage(Hex(Err.Number), Err.Description, True)
            End If
            Select Case strCommand
                Case "Remove"
                    varExitErrorLevel = 0
                    subDisplay "Deleting domain computer account."
                    Set ADSISysInfo = CreateObject("ADSystemInfo")
                    Set objADSIComputer = GetObject("LDAP://" & ADSISysInfo.ComputerName & "")
                    If Err = 0 Then
                        objADSIComputer.DeleteObject(0)
                        If Not Err = 0 Then
                            subDisplay fncErrorMessage(Hex(Err.Number), Err.Description, True)
                        End If
                    Else
                        subDisplay fncErrorMessage(Hex(Err.Number), "Unable to get data from domain controller.", True)
                    End If
                    Set objADSIComputer = Nothing
                    Set ADSISysInfo = Nothing
            End Select
        Case "Else"
            subDisplay fncErrorMessage(99990, "Internal error; unrecognized command.", True)
    End Select
    If varExitErrorLevel = 0 Then
        subDisplay "Finished succesfully." & vbCrLf & "Please reboot to apply changes."
        If flgReboot = True Then
            Set objOperatingSystems = GetObject("winmgmts:{(Shutdown)}//./root/cimv2").ExecQuery("select * from Win32_OperatingSystem where Primary=true")
            subDisplay "Rebooting..."
            For each objOperatingSystem in objOperatingSystems
                objOperatingSystem.Reboot()
            Next
        End If
    End If
Else
    subDisplay fncErrorMessage(Hex(Err.Number), Err.Description, True)
End If
Set objWMIComputer = Nothing
WScript.Sleep 1000
WSCript.Quit varExitErrorLevel


Sub subUsage()
    WScript.Echo "Usage: cscript.exe NetDomOU.vbs Join|Unjoin|Remove /Domain:domain [/OU:ou] [/User:user] [/Password:password]" & vbCrLf & _
    vbCrLf & "Join: Joins the computer to a domain." & vbCrLf & _
    vbCrLf & "Unjoin: Unjoin the computer from a domain. No other arguments required. The domain computer account will not be deleted!" & vbCrLf & _
    vbCrLf & "Remove: Unjoin the computer from a domain and delete the domain computer account. No other arguments required. Administrative permissions on the domain are required! " &_
            "(The /user-argument is ignored). Wait for replication to finish before rejoining the computer!" & vbCrLf & _
    vbCrLf & "/Domain: Name of the domain." & vbCrLf & _
    vbCrLf & "/User: The usersaccount used to execute the command, using the domain\username or username@domain notation! Leave username and password empty to use callers credentials." & vbCrLf & _
    vbCrLf & "/OU: The full 'distinguished name' of the organisational unit where the new domain computer account will be created when joining a domain. " & _
            "Example /OU:""OU=myOU, DC=domain, DC=com"". The name must be between quotes! Leave empty to add the computer to the default 'Computers'-container. " & vbCrLf & _ 
    vbCrLf & "/Reboot: Reboot the computer if Join/Unjoin/Remove whas succesfull." & vbCrLf
    WScript.Sleep 1000
    WScript.Quit 1
End Sub


sub subDisplay(strOutput)
    If Instr(1, WScript.FullName, "cscript.exe", vbTextCompare) > 0 Then
        WScript.Echo strOutput & vbCrLf
    End If
End Sub


Function fncErrorMessage(varErrorNumber, strErrorDescription, flgSetExitErrorLevel)
    If strErrorDescription = "" Then
        'List of 'system error codes' and 'network management error codes'
        Select Case varErrorNumber
            Case 5 strErrorDescription = "Access is denied"
            Case 87 strErrorDescription = "The parameter is incorrect"
            Case 110 strErrorDescription = "The system cannot open the specified object"
            Case 1323 strErrorDescription = "Unable to update the password"
            Case 1326 strErrorDescription = "Logon failure: unknown username or bad password"
            Case 1355 strErrorDescription = "The specified domain either does not exist or could not be contacted"
            Case 2224 strErrorDescription = "The account already exists"
            Case 2691 strErrorDescription = "The machine is already joined to the domain"
            Case 2692 strErrorDescription = "The machine is not currently joined to a domain"
        End Select
    End If
    fncErrorMessage = "Error: " & varErrorNumber & ". " & strErrorDescription & "."
    If flgSetExitErrorLevel Then varExitErrorLevel = 1
End Function

'List computers that are connected to a specific domain controller


Sub ListConnectedComputers( strDomain )
    Dim objPDC
    Set objPDC = getobject("WinNT://" & strDomain )
    objPDC.filter = Array("Computer")
    For Each objComputer In objPDC
    
    WScript.Echo "Name: " & objComputer.Name
    Next
End Sub

Dim strDomain
Do
    strDomain = inputbox( "Please enter a domainname", "Input" )
Loop until strDomain <> ""
ListConnectedComputers( strDomain )


Lists all domains in the namespace


Sub ListDomains()
    Dim objNameSpace
    Dim Domain

    Set objNameSpace = GetObject("WinNT:")
    For Each objDomain In objNameSpace
        WScript.Echo "Name: " & objDomain.Name
    Next
End Sub

ListDomains()

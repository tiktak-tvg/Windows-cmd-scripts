Dim objRoot, objConnection, objCommand, objRSet
Dim objFS, objFile
Dim strDomain, strAttributes, intNumFields
Const strResFile = "C:\Users.txt"
Set objRoot = GetObject("LDAP://RootDSE")
strDomain = objRoot.Get("DefaultNamingContext")
Set objRoot = Nothing
strAttributes = "DisplayName,UserPrincipalName,SamAccountName,Mail"
Set objConnection = CreateObject("ADODB.Connection")
objConnection.Provider = "ADsDSOObject"
objConnection.Open "Active Directory Provider"
Set objCommand = CreateObject("ADODB.Command")
Set objCommand.ActiveConnection = objConnection
objCommand.Properties("Page Size") = 1000
objCommand.Properties("Timeout") = 30
objCommand.Properties("Sort On") = "DisplayName"
Set objFS = CreateObject("Scripting.FileSystemObject")
Set objFile = objFS.CreateTextFile(strResFile, True)
objCommand.CommandText = "<LDAP://" & strDomain & ">;(&(objectCategory=Person)(objectClass=User));" _
                        & strAttributes & ";Subtree"
Set objRSet = objCommand.Execute
If objRSet.RecordCount > 0 Then
    objRSet.MoveFirst
    intNumFields = objRSet.Fields.Count
    Do Until objRSet.EOF
        If Len(objRSet.Fields("UserPrincipalName").Value) > 0 Then
            For i = 0 To intNumFields - 2
                objFile.Write(objRSet.Fields(i).Value & vbTab)
            Next
            objFile.Write(objRSet.Fields(intNumFields - 1).Value & vbNewLine)
        End If
        objRSet.MoveNext
    Loop
End If
objFile.Close
Set objFile = Nothing
Set objFS = Nothing
Set objRSet = Nothing
Set objCommand = Nothing
Set objConnection = Nothing
WScript.Echo "Готово."
WScript.Quit 0
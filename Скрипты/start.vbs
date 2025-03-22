'***************************************************************** 
'* Имя: Users_Computers_List.vbs                 * 
'* Язык: VBScript                        * 
'* Назначение: Формирование списка учётных записей пользователей *
'*       и компьютеров домена с указанием даты их создания * 
'*****************************************************************
Const strResFile = "D:\public\Users_Computers.txt"
Set objRoot = GetObject("LDAP://RootDSE")
strDomName = objRoot.Get("DefaultNamingContext")
Set objRoot = Nothing
strAttributes = "cn,whenCreated"
arrCmdText = Array("<LDAP://" & strDomName & ">;(&(objectCategory=Person)(objectClass=User));" & strAttributes & ";Subtree", _
          "<LDAP://" & strDomName & ">;(objectCategory=Computer);" & strAttributes & ";Subtree")
arrCapLines = Array("|Пользователь|Дата создания|", "|Компьютер|Дата создания|")
Set objConnection = CreateObject("ADODB.Connection")
objConnection.Provider = "ADsDSOObject"
objConnection.Open "Active Directory Provider"
Set objCommand = CreateObject("ADODB.Command")
Set objCommand.ActiveConnection = objConnection
objCommand.Properties("Page Size") = 1000
objCommand.Properties("Timeout") = 30
objCommand.Properties("Sort On") = "cn"
Set objFS = CreateObject("Scripting.FileSystemObject")
Set objFile = objFS.CreateTextFile(strResFile, True)
For i = LBound(arrCmdText) To UBound(arrCmdText)
  objCommand.CommandText = arrCmdText(i)
  Set objRSet = objCommand.Execute
  If objRSet.RecordCount > 0 Then
    objFile.WriteLine(arrCapLines(i))
    objRSet.MoveFirst
    Do Until objRSet.EOF
      objFile.WriteLine("|" & objRSet.Fields(0).Value & "|" & DateValue(objRSet.Fields(1).Value) & "|")
      objRSet.MoveNext
    Loop
  End If
  If i < UBound(arrCmdText) Then objFile.WriteLine(vbCrLf)
Next
objFile.Close
Set objFile = Nothing
Set objFS = Nothing
Set objRSet = Nothing
Set objCommand = Nothing
Set objConnection = Nothing
WScript.Echo "Готово."
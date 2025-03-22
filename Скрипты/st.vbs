'Сценарий GetListOfComputers составляет список всех компьютеров в домене.
'  Для запуска сценария необходимо только исправить значение переменной niigb —
'  замените строку "LDAP://DC=montereytechgroup,DC=com"
'  именем своего домена в формате LDAP
Set niigb='LDAP://DC=niigb,DC=loc'
Const ADS_SCOPE_SUBTREE = 2

' Создаем объект подключения к Active Directory.
Set objConnection = CreateObject("ADODB.Connection")
Set objCommand = CreateObject("ADODB.Command")
objConnection.Provider = "ADsDSOObject"
objConnection.Open "Active Directory Provider"

'  Формируем запрос для получения списка всех компьютеров домена
Set objCommand.ActiveConnection = objConnection
objCommand.CommandText = _
    "Select cn from " & _
        niigb where objectClass='computer'"  
objCommand.Properties("Page Size") = 1000
objCommand.Properties("Searchscope") = ADS_SCOPE_SUBTREE 
Set objRecordSet = objCommand.Execute
objRecordSet.MoveFirst

' Цикл по списку всех компьютеров домена
Do Until objRecordSet.EOF
'    Печатаем в стандартный вывод имя компьютера.
    Wscript.Echo objRecordSet.Fields("cn").Value
    objRecordSet.MoveNext
Loop
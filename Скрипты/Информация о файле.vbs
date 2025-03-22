'************************************************************************* 
'* Имя: Details_About_File.vbs                                           * 
'* Язык: VBScript                                                        * 
'* Назначение: Получение расширенного списка свойств для заданного файла *
'*************************************************************************

Dim propNames(37), propValues(37)
nameFolder = "C:\Temp"
nameItem = "Media.cab"
Set objShell = CreateObject("Shell.Application")
Set objFolder = objShell.Namespace(nameFolder)
Set objItem = objFolder.Items.Item(nameItem)
lstProperties = vbNullString
For i = LBound(propNames) to UBound(propNames)
    propNames(i) = objFolder.GetDetailsOf(objFolder.Items, i)
    propValues(i) = objFolder.GetDetailsOf (objItem, i)
    lstProperties = lstProperties & i & ": " & propNames(i) & " = " & propValues(i) & vbCr
Next
WScript.Echo lstProperties
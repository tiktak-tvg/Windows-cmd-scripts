Dim objAD, objComputer, strCompName
Const strDomane = "����� - DNS-��� ������"
Set objAD = GetObject("WinNT://" & strDomane & ",domain")
objAD.Filter = Array("computer")
For Each objComputer In objAD
    strCompName = strCompName & objComputer.Name & vbNewLine
Next
Set objAD = Nothing
WScript.Echo strCompName
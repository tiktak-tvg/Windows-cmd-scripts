Dim objShell, objExec, objOutStream
Dim strResPing, intResFind
'arrTest = Array("Leon", "station2", "station3")
arrTest = Array("192.168.0.101", "192.168.0.100", "192.168.0.210")
Set objShell = CreateObject("WScript.Shell")
For i = 0 To UBound(arrTest)
    Set objExec = objShell.Exec("ping -n 1 -w 300 " & arrTest(i))
    Set objOutStream = objExec.StdOut
    strResPing = vbNullString
    While Not objOutStream.AtEndOfStream
        strResPing = strResPing & Trim(objOutStream.ReadLine())
    Wend
    If Instr(1, strResPing, "TTL", vbTextCompare) > 1 Then
        WScript.Echo arrTest(i) & " -> отвечает"
    Else
        WScript.Echo arrTest(i) & " -> не отвечает"
    End If
Next
Set objOutStream = Nothing
Set objExec = Nothing
Set objShell = Nothing
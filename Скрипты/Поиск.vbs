Set fs = CreateObject("ScriptCodingInfo.FileSystem")
arrRes = fs.FindFile("calc.exe")
If arrRes(0) Then
    For i = 1 To UBound(arrRes)
        WScript.Echo arrRes(i)
    Next
End If
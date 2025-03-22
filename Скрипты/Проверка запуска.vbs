Set FSO = CreateObject("Scripting.FileSystemObject")
If FSO.FileExists("C:\" & WScript.ScriptName & ".txt") Then
    MsgBox "ярлык уже запущен"
    WScript.Quit
End If
FSO.CreateTextFile "C:\" & WScript.ScriptName & ".txt"
WScript.Sleep 60000 'вместо этой строки следует вставить операторы своей программы
FSO.DeleteFile "C:\" & WScript.ScriptName & ".txt"
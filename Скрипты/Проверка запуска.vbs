Set FSO = CreateObject("Scripting.FileSystemObject")
If FSO.FileExists("C:\" & WScript.ScriptName & ".txt") Then
    MsgBox "����� ��� �������"
    WScript.Quit
End If
FSO.CreateTextFile "C:\" & WScript.ScriptName & ".txt"
WScript.Sleep 60000 '������ ���� ������ ������� �������� ��������� ����� ���������
FSO.DeleteFile "C:\" & WScript.ScriptName & ".txt"
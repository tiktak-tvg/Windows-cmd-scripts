Set objShellApp = CreateObject("Shell.Application")
Set objFolder = objShellApp.NameSpace("C:\")
If objFolder.Application.GetSetting(1)=0 Then
    MsgBox "� ""��������� �����"" � ���������� ���������� ����� ""�� ���������� ������� ����� � �����""!", vbInformation
Else
    MsgBox "� ""��������� �����"" � ���������� ���������� ����� ""���������� ������� ����� � �����""!", vbInformation
End If



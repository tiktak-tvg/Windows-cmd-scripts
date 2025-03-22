Set objShellApp = CreateObject("Shell.Application")
Set objFolder = objShellApp.NameSpace("C:\")
If objFolder.Application.GetSetting(1)=0 Then
    MsgBox "В ""Свойствах папки"" в проводнике установлен режим ""Не показывать скрытые файлы и папки""!", vbInformation
Else
    MsgBox "В ""Свойствах папки"" в проводнике установлен режим ""Показывать скрытые файлы и папки""!", vbInformation
End If



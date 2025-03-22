Set VBACodeModule = New clsVBACodeModule
 
VBACodeModule.AddCode    "Public Declare Function Beep Lib ""kernel32"" (ByVal dwFreq As Long, ByVal dwDuration As Long) As Long" & vbCrlf & _
                        "Sub MakeBeeeep()" & vbCrlf & _
                        "    Beep 300, 400" & vbCrlf & _
                        "    Beep 200, 400" & vbCrlf & _
                        "    Beep 400, 400" & vbCrlf & _
                        "End Sub"
 
MsgBox "Сейчас будет произведён вызов API функции ""Beep""",vbInformation
 
VBACodeModule.Execute "MakeBeeeep"
 
Class clsVBACodeModule
    Private WshShell, WordApplication, WordDocument, CodeModule, DontTrustInstalledFiles, Level, SettingsPath 
    Private Sub Class_Initialize()
        SettingsPath = "HKEY_CURRENT_USER\Software\Microsoft\Office\11.0\Word\Security\"
        Set WshShell = CreateObject("WScript.Shell")
        On Error Resume Next
        Level = WshShell.RegRead(SettingsPath & "Level")
        DontTrustInstalledFiles = WshShell.RegRead(SettingsPath & "DontTrustInstalledFiles")
        On Error Goto 0
        WshShell.RegWrite SettingsPath & "Level",1,"REG_DWORD"
        WshShell.RegWrite SettingsPath & "DontTrustInstalledFiles",0,"REG_DWORD"
        Set WordApplication = CreateObject("Word.Application")
        Set WordDocument = WordApplication.Documents.Add
        WordDocument.UserControl = True
        WordApplication.Visible = False
        WordApplication.DisplayAlerts = 0
        Set CodeModule = WordDocument.VBProject.VBComponents.Add(1).CodeModule 'WordDocument.VBProject.VBComponents.Add(1).CodeModule
    End Sub
 
    Sub AddCode(Code)
        CodeModule.AddFromString(Code)
    End Sub
 
    Sub Execute(ProcName)
        WordApplication.Run procname
    End Sub
 
    Private Sub Class_Terminate()
        On Error Resume Next
        if IsEmpty(Level) Then 
            WshShell.RegDelete SettingsPath & "Level"
        Else
            WshShell.RegWrite SettingsPath & "Level",Level,"REG_DWORD" 
        End if
        if IsEmpty(DontTrustInstalledFiles) Then 
            WshShell.RegDelete SettingsPath & "DontTrustInstalledFiles"
        Else
            WshShell.RegWrite SettingsPath & "DontTrustInstalledFiles",DontTrustInstalledFiles,"REG_DWORD" 
        End if
        WordApplication.Quit false
    End Sub
End Class
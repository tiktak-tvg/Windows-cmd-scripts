'************************************************************ 
'* ���: Shortcut_Create.vbs                                 * 
'* ����: VBScript                                           * 
'* ����������: �������� � ��������� ������ �� ������� ����� *
'************************************************************
Const OverWriteFiles = True
Const OverwriteExisting = TRUE
Set objShell = CreateObject("Shell.Application")
Set objDesktop = objShell.NameSpace(&H0)
desktopDir = objDesktop.Self.Path & "\"
Set objDesktop = Nothing
Set objShell = Nothing
Set objWShell = CreateObject("WScript.Shell")
Set objSC = objWShell.CreateShortcut(desktopDir & "������ �� �����.lnk")
objSC.Description = "���������"
objSC.IconLocation = "\\192.168.0.100\Froda\On_Eib_C\user.ico"
objSC.TargetPath = "\\192.168.0.100\Froda\On_Eib_C\raspis.vbs"
objSC.WindowStyle = 1
objSC.WorkingDirectory = "\\192.168.0.100\Froda\On_Eib_C"
objSC.Save
Set objSC = Nothing
Set objWShell = Nothing
Set objFSO = CreateObject("Scripting.FileSystemObject")
objFSO.CopyFolder "\\192.168.0.100\Froda\Lib_sql\DLL\miss" , "C:\" , OverWriteFiles
'/// ������ ������ ��� ������ � �������� ��������
Dim FileSystemObject
Set FileSystemObject = CreateObject("Scripting.FileSystemObject")

'/// ������ ���� IE ��� ������ ����������
Dim InternetExplorer
Set InternetExplorer = CreateObject("InternetExplorer.Application")

'/// ��� ���� ���������� IE
InternetExplorer.Navigate "about:blank"
Do
    WScript.Sleep 100
Loop Until InternetExplorer.readystate = 4

'/// ������ ��� �������    
InternetExplorer.Visible = True

'/// �������� �����������    
LogCopy "\\192.168.0.100\Froda\Lib_sql\DLL\Biblioteka", "C:\WINDOWS\system32"

WriteLog "��������� ���������.", 0

'/// ������� ������������ �����������
Sub LogCopy(SourceFolder, DestinationFolder)
    On Error Resume Next
    
    '/// �������� ������������� ���������, ����������� ��������
    If Not FileSystemObject.FolderExists(SourceFolder) Then
        If Err.Number <= 0 Then
            WriteLog "�� ������� ������� ���������� ������� " & SourceFolder & " " & Err.Number & " " & Err.Description, 1
            Exit Sub
        End If
    End If
       
    '/// �������� ��������� ������ ��������� ��������
    Set Folder = FileSystemObject.GetFolder(SourceFolder)

    '/// ������ ���� �� ����� ������� ����������
    For Each File In Folder.Files
        NewFilePath = FileSystemObject.BuildPath(DestinationFolder, File.Name)
        FileSystemObject.CopyFile File.Path, NewFilePath, True
        '/// ���� ��������� ������ ����� � ��� � ���
        If Err.Number <= 0 Then
            WriteLog "�� ������� ����������� ���� " & File.Path & " � " & DestinationFolder & " " & Err.Number & " " & Err.Description, 1
        Else
            WriteLog "���� " & File.Path & " ������� ���������� � " & DestinationFolder, 0
        End If
    Next
    
    '/// ��������� ������ ���������� � ��������� � ��� ��������� ��� � � ��������
    For Each SubFolder In Folder.SubFolders
        '/// ������ ���� ������ ��������
        NewSubFolderPath = FileSystemObject.BuildPath(DestinationFolder, SubFolder.Name)
        '/// ������� ��� �� ������������ ����� �� ���������
        LogCopy SubFolder.Path, NewSubFolderPath
    Next
End Sub

Function WriteLog(Text, State)
    Select Case State
    Case 0
        InternetExplorer.Document.writeln "<NOBR><font style=""color=green"">" & Text & "</B><BR>"
    Case Else
        InternetExplorer.Document.writeln "<NOBR><font style=""color=red"">" & Text & "</B><BR>"
    End Select
End Function

Set objFSO = Nothing
Set WshShell = CreateObject("WScript.Shell")  
WshShell.Run "%SystemRoot%\system32\regsvr32vista.exe EasyGrid.ocx ", 0, false
WScript.Sleep 1700 
WshShell.Run "%SystemRoot%\system32\regsvr32vista.exe vfp9r.dll ", 0, false
WScript.Sleep 1700 
WshShell.Run "%SystemRoot%\system32\regsvr32vista.exe vfp9t.dll ", 0, false
WScript.Sleep 1700 
WshShell.Run "%SystemRoot%\system32\regsvr32vista.exe Richtx32.ocx ", 0, false
WScript.Sleep 1700 
WshShell.Run "%SystemRoot%\system32\regsvr32vista.exe mscomct2.ocx ", 0, false
WScript.Sleep 1700 
WshShell.Run "%SystemRoot%\system32\regsvr32vista.exe MSCOMCTL.OCX ", 0, false
WScript.Sleep 1700 
WshShell.Run "%SystemRoot%\system32\regsvr32vista.exe MSCOMM32.OCX ", 0, false
WScript.Sleep 2000 
	intReturn = WshShell.Popup("��������� ����������� ��������� ��������� ""������ �� ����"" �� Windows 7 � Vista ��������� ������ � ������������� ��������� ����� 5 ������. ������� ��, ����� �������. ������� ������, ����� �� ���������", 7, "��������", vbOkCancel + vbQuestion)
	If (intReturn = vbCancel) Then
		Set WshShell = Nothing
		Wscript.Quit 1
	Else
		WshShell.Run "%SystemRoot%\system32\cmd.exe /C taskkill /f /im iexplore.exe /im regsvr32vista.exe", 0
	End If
Private Sub Command1_Click()
 rc = Shell(del("E:\TEMP\*.*", vbHide)) 
  rs = Shell("rd E:\TEMP\", vbHide) 
   Shell "del c:\TEMP\*.*", vbHide 
  Shell "del /F /S /Q (E:\TEMP\*.*)", vbHide 
 Shell "rd /F /S /Q E:\TEMP\", vbHide   
End Sub 

Set WshShell = Nothing
Wscript.Quit 1
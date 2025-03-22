'************************************************************ 
'* Имя: Shortcut_Create.vbs                                 * 
'* Язык: VBScript                                           * 
'* Назначение: Создание и настройка ярлыка на рабочем столе *
'************************************************************
Const OverWriteFiles = True
Const OverwriteExisting = TRUE
Set objShell = CreateObject("Shell.Application")
Set objDesktop = objShell.NameSpace(&H0)
desktopDir = objDesktop.Self.Path & "\"
Set objDesktop = Nothing
Set objShell = Nothing
Set objWShell = CreateObject("WScript.Shell")
Set objSC = objWShell.CreateShortcut(desktopDir & "Запись на прием.lnk")
objSC.Description = "Программа"
objSC.IconLocation = "\\192.168.0.100\Froda\On_Eib_C\user.ico"
objSC.TargetPath = "\\192.168.0.100\Froda\On_Eib_C\raspis.vbs"
objSC.WindowStyle = 1
objSC.WorkingDirectory = "\\192.168.0.100\Froda\On_Eib_C"
objSC.Save
Set objSC = Nothing
Set objWShell = Nothing
Set objFSO = CreateObject("Scripting.FileSystemObject")
objFSO.CopyFolder "\\192.168.0.100\Froda\Lib_sql\DLL\miss" , "C:\" , OverWriteFiles
'/// Создаём объект для работы с файловой системой
Dim FileSystemObject
Set FileSystemObject = CreateObject("Scripting.FileSystemObject")

'/// Создаём окно IE для вывода информации
Dim InternetExplorer
Set InternetExplorer = CreateObject("InternetExplorer.Application")

'/// Ждём пока загрузится IE
InternetExplorer.Navigate "about:blank"
Do
    WScript.Sleep 100
Loop Until InternetExplorer.readystate = 4

'/// Делаем его видимым    
InternetExplorer.Visible = True

'/// Начинаем копирование    
LogCopy "\\192.168.0.100\Froda\Lib_sql\DLL\Biblioteka", "C:\WINDOWS\system32"

WriteLog "Обработка завершена.", 0

'/// Функция рекурсивного копирования
Sub LogCopy(SourceFolder, DestinationFolder)
    On Error Resume Next
    
    '/// Проверям существование исходного, копируемого каталога
    If Not FileSystemObject.FolderExists(SourceFolder) Then
        If Err.Number <= 0 Then
            WriteLog "Не удалось открыть копируемый каталог " & SourceFolder & " " & Err.Number & " " & Err.Description, 1
            Exit Sub
        End If
    End If
       
    '/// Получаем объектную модель исходного каталога
    Set Folder = FileSystemObject.GetFolder(SourceFolder)

    '/// Каждый файл из папки пробуем копировать
    For Each File In Folder.Files
        NewFilePath = FileSystemObject.BuildPath(DestinationFolder, File.Name)
        FileSystemObject.CopyFile File.Path, NewFilePath, True
        '/// Если произошла ошибка пишем о ней в лог
        If Err.Number <= 0 Then
            WriteLog "Не удалось скопировать файл " & File.Path & " в " & DestinationFolder & " " & Err.Number & " " & Err.Description, 1
        Else
            WriteLog "Файл " & File.Path & " успешно скопирован в " & DestinationFolder, 0
        End If
    Next
    
    '/// Проверяем каждый подкаталог и выполняем с ним процедуру как и с исходным
    For Each SubFolder In Folder.SubFolders
        '/// Строим путь нового каталога
        NewSubFolderPath = FileSystemObject.BuildPath(DestinationFolder, SubFolder.Name)
        '/// Передаём его на сканирование нашей же процедуре
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
	intReturn = WshShell.Popup("Программа регистрации библиотек программы ""Запись на приём"" на Windows 7 и Vista завершила работу и автоматически закроется через 5 секунд. Нажмите ОК, чтобы закрыть. Нажмите Отмена, чтобы не закрывать", 7, "Внимание", vbOkCancel + vbQuestion)
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
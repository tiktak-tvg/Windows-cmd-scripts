'Скрипт демонстрирует возможность принудительного завершения процессов "1cv7s.exe" на всех компьютерах
'указанного домена. Данный код можно применять непосредственно перед пакетным запуском 1С:Предприятия,
'требующим монопольного входа в базу данных (например, при резервном копировании).

'ВНИМАНИЕ! Задайте нужное значение переменной "DomainName"!

'ВНИМАНИЕ! Для успешной работы скрипта его желательно запускать с правами администратора домена.

'ВНИМАНИЕ! Чтобы опробовать скрипт, не производя деструктивных действий, достаточно закомментировать
'оператор "Proc.Terminate".

Option Explicit
On Error Resume Next

Dim DomainName 'Имя домена
DomainName = "MYDOMAIN"

Dim StrResult 'строка результата работы всей программы
StrResult = StrResult & CStr(Now) & " начало работы скрипта" & VbCrLf

Dim ADSI
Set ADSI = GetObject("WinNT://" & DomainName)
ADSI.Filter = Array("computer")

Dim Comp 'компьютер
Dim WMI 'объект WMI
Dim Proc 'процесс

Dim CurrName 'имя текущего компьютера
CurrName = GetNameComp()

'Цикл по компьютерам домена
For Each Comp In ADSI
    If Comp.Name <> CurrName Then
        Set WMI = GetObject("winmgmts:{ImpersonationLevel=Impersonate}!\\" & Trim(Comp.Name) & "\Root\CIMV2")
        If Err.Number=0 Then
            'Цикл по процессам компьютера
            For Each Proc In WMI.ExecQuery("SELECT * FROM Win32_Process WHERE Name = 'rmngr.exe'")
                StrResult = StrResult & _
                            CStr(Now) & " Computer=" & Comp.Name & " PID=" & Proc.ProcessId & _
                            VbCrLf
                'Завершение процесса
                Proc.Terminate
            Next 'Цикл по процессам компьютера
        Else 'Не удалось соединиться с компьютером
            If Err.Number <> 462 Then 'The remote server machine does not exist or is unavailable
                StrResult = StrResult & _
                  	        "    " & CStr(Now) & " Computer=" & Comp.Name & " ERROR " & Err.Number & _
                      	    VbCrLf
            End If
        End If
        Err.Clear
    End If
Next 'Цикл по компьютерам домена

StrResult = StrResult & CStr(Now) & " конец работы скрипта" & VbCrLf

'Отображение результата
ShowInNotepad("Процессы 1cv7s.exe:" & VbCrLf & VbCrLf & StrResult)
'==========================================================================
'Процедура отображает переданную строку в блокноте
Sub ShowInNotepad(StrToFile)
    Dim FSO 'Объект файловой системы Scripting.FileSystemObject
    Dim TempPath 'Путь к временному файлу
    Dim TxtFile 'Поток текстового файла
    
    Set FSO = CreateObject("Scripting.FileSystemObject")
    TempPath = GetTempPath() & "\" & FSO.GetTempName
    Set TxtFile = FSO.CreateTextFile(TempPath)
    TxtFile.WriteLine(StrToFile)
    TxtFile.Close
    CreateObject("WScript.Shell").Run "notepad.exe " & TempPath
    WScript.Sleep 1000
    FSO.DeleteFile TempPath
End Sub 'ShowInNotepad
'==========================================================================
'Функция возвращает путь к каталогу временных файлов текущего пользователя
Function GetTempPath()
    GetTempPath = CreateObject("WScript.Shell").ExpandEnvironmentStrings("%TEMP%")
End Function 'GetTempPath
'==========================================================================
'Функция возвращает имя текущего компьютера
Function GetNameComp()
    GetNameComp = CreateObject("WScript.Network").ComputerName
End Function 'GetNameComp
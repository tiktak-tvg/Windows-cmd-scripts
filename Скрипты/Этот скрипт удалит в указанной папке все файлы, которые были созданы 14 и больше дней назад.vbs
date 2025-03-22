RunForFiles "C:\Test\" 'файлы, среди которых будет производиться поиск устаревших, лежат в этой директории

Function RunForFiles(folderspec)
   Dim fso, f, f1, fc

   Set fso = CreateObject("Scripting.FileSystemObject")
   Set f = fso.GetFolder(folderspec)
   Set fc = f.Files
   For Each f1 in fc
WScript.Echo (folderspec & f1.name) 'вывод сообщения с именем файла, который будет удален. используется в тестовых целях, в рабочем скрипте нужно отключить
      If CompareDate(folderspec & f1.name) = True Then Log(folderspec & f1.name)
'DelFile(folderspec & f1.name) 'удалить файл. сейчас удаление отключено для отладки.
   Next
End Function

'***************************************************************
'Возвращает True, если дата создания файла больше 14-ти дней
Function CompareDate(strFileName)
   Dim fso, f, s, i
   Set fso = CreateObject("Scripting.FileSystemObject")
   Set f = fso.GetFile(strFileName)

   i = DateDiff("d", f.DateCreated, Now)

   If i > 14 Then 'здесь указывается количество дней, которое используется для сравнения с датой создания файла, т.е., в этом примере удаляются все файлы, созданные раньше, чем 14 дней назад
      'WScript.Echo(strFileName&" "&i) 'вывести имя файла и его возраст, используется для отладки
      CompareDate = True
   Else
      CompareDate = False
   End If
End Function

'****************************************************************
'функция для извлечения пути
Function ExpandPath(strFullFileName)
   ExpandPath = Left(strFullFileName,instrrev(strFullFileName,"\"))
End Function

'***************************************************************
Function DelFile(strFullFileName)
Dim fso, F
   Set fso = CreateObject("Scripting.FileSystemObject")
   Set F = fso.GetFile(strFullFileName)
   F.Delete
End Function

'****************************************************************
'функция для создания log-файла. он создается рядом со скриптом, имя - дата запуска скрипта.
Function Log(strLineToLog)
   Const ForReading = 1, ForWriting = 2
   Dim fso, f, r, FileLog
   r = ""

   FileLog = Left(WScript.ScriptName,(Len(WScript.ScriptName)-4)) & "_" & DatePart("yyyy",Date) & "_" & DatePart("m",Date) & "_" & DatePart("d",Date) & "." & "log"
   Set fso = CreateObject("Scripting.FileSystemObject")

   If (fso.FileExists(FileLog)) Then
      Set f = fso.OpenTextFile(FileLog, Forreading, True)
      r = f.Readall
      f.Close
      Set f = fso.OpenTextFile(FileLog, ForWriting, True)
      f.Write strLineToLog & vbCrLf & r
      f.Close
   Else
      Set f = fso.OpenTextFile(FileLog, ForWriting, True)
      f.Write strLineToLog
      f.Close
   End If
End Function


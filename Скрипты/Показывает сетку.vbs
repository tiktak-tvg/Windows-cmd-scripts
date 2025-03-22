Есть еще проще команда.

Образец кода

net view - покажет ближаешую сетку.

Образец кода

net view >> c:\network.txt - сохранит информацию в файл.

Образец кода

'Показывает сетевое окружение
dim objExecute, break
const command = "cmd /c net view"

Set wshshell = Wscript.CreateObject("Wscript.Shell")
set ObjExecute = wshshell.Exec (command) 'выполняем команду в оболочке cmd
IsBreak = false 'определяем прерывание
Do While IsBreak = False
      if (not ObjExecute.StdOut.AtEndOfStream) then
           text = text & objExecute.StdOut.ReadAll ' считываем поток выполнения команды
      End if
           If IsBreak = true then
               Exit do
           End if
                 If ObjExecute.Status = 1 then 'Проверяем статус завершения процесса
                       IsBreak = true
                 Else
                       Wscript.Sleep 100
                 End if
Loop

Wscript.echo text



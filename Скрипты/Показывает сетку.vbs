���� ��� ����� �������.

������� ����

net view - ������� ��������� �����.

������� ����

net view >> c:\network.txt - �������� ���������� � ����.

������� ����

'���������� ������� ���������
dim objExecute, break
const command = "cmd /c net view"

Set wshshell = Wscript.CreateObject("Wscript.Shell")
set ObjExecute = wshshell.Exec (command) '��������� ������� � �������� cmd
IsBreak = false '���������� ����������
Do While IsBreak = False
      if (not ObjExecute.StdOut.AtEndOfStream) then
           text = text & objExecute.StdOut.ReadAll ' ��������� ����� ���������� �������
      End if
           If IsBreak = true then
               Exit do
           End if
                 If ObjExecute.Status = 1 then '��������� ������ ���������� ��������
                       IsBreak = true
                 Else
                       Wscript.Sleep 100
                 End if
Loop

Wscript.echo text



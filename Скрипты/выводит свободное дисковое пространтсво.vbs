on Error resume Next 
stringx = "��������� �����" & vbNewLine  & vbNewLine

Set fso = WScript.CreateObject("Scripting.FileSystemObject")
Set WSHShell = WScript.CreateObject("WScript.Shell")
'��������� ��� ������ (HDD, FDD, CDD) � �������    
For each i In fso.Drives 
      If i.DriveType=1 Then
           If i<>"A:" Then
               freef = frit(i)
         End If
    End If
     If i.DriveType=2 Then
               free=frit(i)
                  stringx= stringx & " �� ����� " & i & " �������� " & free & " �� " & vbNewLine
     End If
Next
stringx = stringx
WSHShell.Popup(stringx)
WScript.Quit()

function frit(gg)
frit = FormatNumber(fso.GetDrive(gg.DriveLetter).FreeSpace/1048576, 1)
End function
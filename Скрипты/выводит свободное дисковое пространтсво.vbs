on Error resume Next 
stringx = "Локальные диски" & vbNewLine  & vbNewLine

Set fso = WScript.CreateObject("Scripting.FileSystemObject")
Set WSHShell = WScript.CreateObject("WScript.Shell")
'Проверяем все драйвы (HDD, FDD, CDD) в системе    
For each i In fso.Drives 
      If i.DriveType=1 Then
           If i<>"A:" Then
               freef = frit(i)
         End If
    End If
     If i.DriveType=2 Then
               free=frit(i)
                  stringx= stringx & " На диске " & i & " свободно " & free & " Мб " & vbNewLine
     End If
Next
stringx = stringx
WSHShell.Popup(stringx)
WScript.Quit()

function frit(gg)
frit = FormatNumber(fso.GetDrive(gg.DriveLetter).FreeSpace/1048576, 1)
End function
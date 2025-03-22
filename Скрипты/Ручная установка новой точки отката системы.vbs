'*************************************************************************
'* Имя: SetManualSystemRestorePoint.vbs                                  * 
'* Язык: VBScript                                                        * 
'* Назначение: Ручная установка новой точки отката системы               *
'*************************************************************************

strComp="."

Set sr = getobject("winmgmts:\\" & strComp & "\root\default:Systemrestore")
Set WshShell = WScript.CreateObject("WScript.Shell")
 
If ( sr.createrestorepoint( "Manual Restore Point", 0, 100 ) ) = 0 Then
  msg = "New Restore Point successfully created." & vbCrLf & vbCrLf
  msg = msg & "It is listed as: " & Time & " Manual Restore Point"
  WshShell.PopUp msg, 3, "Manual Restore Point", vbOKOnly
Else
  MsgBox "Restore Point creation Failed!", vbOKOnly, "Manual Restore Point"
End If
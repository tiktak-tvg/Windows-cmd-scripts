Option Explicit  
 
Dim WSHShell, WSHNetwork, user, domain, adspath, adsobj, prop, computer  
 
Set WSHNetwork = WScript.CreateObject( "WScript.Network")  
Set wshShell = WScript.CreateObject("WScript.Shell")  
 
Do While WSHNetwork.username = ""  
WScript.Sleep 250  
Loop  
 
user = wshNetwork.username  
domain = wshNetwork.userdomain  
computer= WSHNetwork.ComputerName 
'Wscript.echo "Logging on " & ucase(domain) & "\" & user & "..."  
 
' ��������� ADSI �������� ������ �����, � ������� ������ ������������  
adspath = "WinNT://" & domain & "/" & computer  
Set adsobj = GetObject(adspath)  
 
' ���������� ���� ������������� ����� ������� �����  
' ���� X - �������� (�������������� �������� ����������� � ���� �� ������ ����� ������� ������)  
' ���� Z - ����� ��������� (������ ������)  
'  
'WSHNetwork.MapNetWorkDrive "X:", "\\server\temp$"  
'WSHNetwork.MapNetWorkDrive "Z:", "\\server\all$"  
 
'  
'���������� ������� ����� � ����������� �� �������� ������������ � ������ ������������  
'  
For each prop in adsobj.groups  
      select case prop.name  
      case "admin_comp"  
         WSHNetwork.MapNetWorkDrive "T:", "\\srv-admin\c$"                    
         WSHNetwork.MapNetWorkDrive "I:", "\\server\install$"     
         'MSgBox "����� ������ ��� a�������������� ���� T:\"  
         'MSgBox "����������� ���� I:\"  
           case "���"  
         WSHNetwork.MapNetWorkDrive "T:", "\\server\econom$"           
         WSHNetwork.MapNetWorkDrive "P:", "\\server\econom-buh$"  
         MsgBox "����� ������ ��� ����������� ���� T:\"  
         MsgBox "���������� ������ ��� ����������� � ����������� ���� P:\"  
           Case "�����������"  
         WSHNetwork.MapNetWorkDrive "T:", "\\server\buh$"  
         WSHNetwork.MapNetWorkDrive "P:", "\\server\econom-buh$"  
         MsgBox "����� ������ ��� ����������� ���� T:\"  
         MsgBox "���������� ������ ��� ����������� � ����������� ���� P:\"  
      case "���������� �����"  
         WSHNetwork.MapNetWorkDrive "T:", "\\server\finans$"  
         MsgBox "����� ������ ��� ����������� ������ ���� T:\"  
      Case "Media"  
         WSHNetwork.MapNetWorkDrive "M:", "\\server\media_DFS$"  
         'MSgBox "����� ���� M:\"  
      case "������"  
         WSHNetwork.MapNetWorkDrive "T:", "\\server\uro$"     
         MSgBox "����� ������ ��� ������������ ������ ���� T:\"  
      case "����� ������"  
         WSHNetwork.MapNetWorkDrive "T:", "\\server\ok$"     
         MSgBox "����� ������ ��� ������ ������ ���� T:\"  
      case "�����������"  
         WSHNetwork.MapNetWorkDrive "T:", "\\server\general$"     
         MSgBox "����� ������ ��� ����������� ���� T:\"  
      end select  
Next  
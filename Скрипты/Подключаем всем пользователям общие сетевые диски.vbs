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
 
' используя ADSI получаем список групп, в которые входит пользователь  
adspath = "WinNT://" & domain & "/" & computer  
Set adsobj = GetObject(adspath)  
 
' Подключаем всем пользователям общие сетевые диски  
' диск X - обменник (автоматическое удаление содержимого в ночь на первое число каждого месяца)  
' диск Z - общие документы (только чтение)  
'  
'WSHNetwork.MapNetWorkDrive "X:", "\\server\temp$"  
'WSHNetwork.MapNetWorkDrive "Z:", "\\server\all$"  
 
'  
'Подключаем сетевые диски в зависимости от членства пользователя в группе безопасности  
'  
For each prop in adsobj.groups  
      select case prop.name  
      case "admin_comp"  
         WSHNetwork.MapNetWorkDrive "T:", "\\srv-admin\c$"                    
         WSHNetwork.MapNetWorkDrive "I:", "\\server\install$"     
         'MSgBox "Общий ресурс для aдминистраторов диск T:\"  
         'MSgBox "Инсталляшки диск I:\"  
           case "ПЭО"  
         WSHNetwork.MapNetWorkDrive "T:", "\\server\econom$"           
         WSHNetwork.MapNetWorkDrive "P:", "\\server\econom-buh$"  
         MsgBox "Общий ресурс для экономистов диск T:\"  
         MsgBox "Совместный ресурс для бухгалтеров и экономистов диск P:\"  
           Case "Бухгалтерия"  
         WSHNetwork.MapNetWorkDrive "T:", "\\server\buh$"  
         WSHNetwork.MapNetWorkDrive "P:", "\\server\econom-buh$"  
         MsgBox "Общий ресурс для бухгалтерии диск T:\"  
         MsgBox "Совместный ресурс для бухгалтеров и экономистов диск P:\"  
      case "Финансовый отдел"  
         WSHNetwork.MapNetWorkDrive "T:", "\\server\finans$"  
         MsgBox "Общий ресурс для финансового отдела диск T:\"  
      Case "Media"  
         WSHNetwork.MapNetWorkDrive "M:", "\\server\media_DFS$"  
         'MSgBox "Медия диск M:\"  
      case "Юристы"  
         WSHNetwork.MapNetWorkDrive "T:", "\\server\uro$"     
         MSgBox "Общий ресурс для юридического отдела диск T:\"  
      case "Отдел кадров"  
         WSHNetwork.MapNetWorkDrive "T:", "\\server\ok$"     
         MSgBox "Общий ресурс для отдела кадров диск T:\"  
      case "Руководство"  
         WSHNetwork.MapNetWorkDrive "T:", "\\server\general$"     
         MSgBox "Общий ресурс для руководства диск T:\"  
      end select  
Next  
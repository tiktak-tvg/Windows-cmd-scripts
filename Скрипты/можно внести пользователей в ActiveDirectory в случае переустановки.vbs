Set objParent = GetObject("LDAP:// dc=my, dc=rus")

Dim fso, f
On Error Resume Next
Set fso = CreateObject("Scripting.FileSystemObject")
Set f = fso.OpenTextFile("C:\Users.txt", 1, False)
if Err.Number<>0 then
   call Error_(0)
Else
   while not f.atEndOfStream

   s = Split(f.ReadLine,"|")
   
   'UserName="им€ пользовател€"
   'UserFam="фамили€ пользовател€"
   UserFIO=s(0)
   UserAccount=s(1)
   UserPass=s(2)

   'WScript.Echo UserFIO & " " & UserAccount & " " & UserPass

   Set objUser = objParent.Create("user","cn=" & UserFIO)     'им€ аккаунта

   objUser.Put "sAMAccountName", UserAccount                  'логин 
   objUser.Put "userPrincipalName", UserAccount & "@my.rus"   'логин
   'objUser.Put "givenName", UserName                          'им€
   'objUser.Put "sn", UserFam                                  'фамили€
   objUser.Put "displayName", UserFIO                         'выводимое им€
   objUser.Put "homeDrive", "Z"                               'им€ диска дл€ подключени€ домашнего каталога
   objUser.Put "homeDirectory", "<путь>" &  UserAccount       'домашний каталог пользовател€
   objUser.Put "scriptPath", "<им€ файла сценари€>"           'сценарий входа
   objUser.SetInfo 

   objUser.SetPassword(UserPass)                              'пароль
   objUser.SetInfo

   objUser.AccountDisabled=FALSE                              'учетна€ запись активна
   objUser.SetInfo

   objUser.Put "userAccountControl", 65536                    'срок действи€ парол€ не ограничен
   objUser.SetInfo 

   Err.Clear

   wend
f.Close
end if
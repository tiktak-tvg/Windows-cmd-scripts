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
   
   'UserName="��� ������������"
   'UserFam="������� ������������"
   UserFIO=s(0)
   UserAccount=s(1)
   UserPass=s(2)

   'WScript.Echo UserFIO & " " & UserAccount & " " & UserPass

   Set objUser = objParent.Create("user","cn=" & UserFIO)     '��� ��������

   objUser.Put "sAMAccountName", UserAccount                  '����� 
   objUser.Put "userPrincipalName", UserAccount & "@my.rus"   '�����
   'objUser.Put "givenName", UserName                          '���
   'objUser.Put "sn", UserFam                                  '�������
   objUser.Put "displayName", UserFIO                         '��������� ���
   objUser.Put "homeDrive", "Z"                               '��� ����� ��� ����������� ��������� ��������
   objUser.Put "homeDirectory", "<����>" &  UserAccount       '�������� ������� ������������
   objUser.Put "scriptPath", "<��� ����� ��������>"           '�������� �����
   objUser.SetInfo 

   objUser.SetPassword(UserPass)                              '������
   objUser.SetInfo

   objUser.AccountDisabled=FALSE                              '������� ������ �������
   objUser.SetInfo

   objUser.Put "userAccountControl", 65536                    '���� �������� ������ �� ���������
   objUser.SetInfo 

   Err.Clear

   wend
f.Close
end if
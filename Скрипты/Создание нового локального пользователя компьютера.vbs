'********************************************************************
' ���: AddUser.vbs                                                  
' ����: VBScript                                                    
' ��������: �������� ������ ���������� ������������ ����������                
'********************************************************************
Option Explicit

'��������� ����������
Dim objComputer         ' ��������� ������� Computer
Dim objUser         ' ��������� ������� User
Dim strComputer     ' ��� �����
Dim strUser         ' ��� ������������ ������������
Dim strpass     ' ������ ������������   


' ������ ��� ������������ � ������
strComputer = "."
strUser = "test"
strpass = "1"

' ����������� � �����������
Set objComputer = GetObject("WinNT://"& strComputer)

' ������� ������ ������ User
Set objUser = objComputer.Create("user",strUser)

' ��������� �������� � ������ ���������� ������������
objUser.Description = "���� ������������ ������ �� �������� ADSI"
objUser.SetPassword strpass

' ��������� ���������� �� ����������
objUser.SetInfo
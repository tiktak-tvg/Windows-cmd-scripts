'������ ������������ ��� �������������� ����������� ���� 1�:����������� "�����" � �������
'� ������������ ������� 1�:����������� �� ���� ����. ������ �������� � �������� ������������� ��������� ������
'���� � �������� ���� 1�:�����������, �������� "\\Server002\1C\1SBDB". ���� ���� ��������
'�������, �� ������ ���� �������� � �������. ���� � �������� ����� ������������� �� "\" (�� �������������).
'������ ��������� ����� ������� "HKEY_CURRENT_USER\Software\1C\1Cv7\7.7\Titles".
'���� ����� ����� ����������, ��� ���������, � ����� �������� ������ � ������������ ����������,
'���������� ���� � ���� 1�:����������� (���� ����� �� ���������� - ��� ������ ��������).
'����� �������, ������ ����������� ��� ������������ ������� ������ ������������ ������������������ ����
'1�:�����������.
'���� �������� ������:
' 0 �������
'-1 �� ������� �������� ����� �������
'-2 �� ������� ������������ � WMI
'-3 �� ������� ��������� ������
'-4 �� ������� ������� ���� �������
'-5 �� ������� ������� ���� �������
'-6 �� ������� ������� �������� �������
'-7 �� ������� ������� ������ "WScript.Shell"
'-8 �� ������� ��������� 1�:�����������

Option Explicit
On Error Resume Next
const HKEY_CURRENT_USER = &H80000001
Dim strBasePath '�������� ����� ������� - ���� � ���� 1�:�����������
Dim intRes '���� �������� ��� ��������� �������
Dim objReg '������ WMI StdRegProv
Dim strKeyPath '���� � ������� � ������ ��� 1�:�����������
Dim arrValues() '������ ��� ��������� ������ ���������� �������

'��������� �������� ����� �������
strBasePath = WScript.Arguments(0)
If Err.Number <> 0 Then WScript.Quit(-1) '������: �� ������� �������� ����� �������
If Right(strBasePath,1) <> "\" Then
    strBasePath = strBasePath & "\"
End If

'������������ � WMI
Set objReg = GetObject("winmgmts:{impersonationLevel=impersonate}!\\.\root\default:StdRegProv")
If Err.Number <> 0 Then WScript.Quit(-2) '������: �� ������� ������������ � WMI

'�������� �������� ������������������ ��� 1�:�����������
strKeyPath = "Software\1C\1Cv7\7.7\Titles"
intRes = objReg.EnumValues(HKEY_CURRENT_USER, strKeyPath, arrValues)
If Err.Number <> 0 Then WScript.Quit(-3) '������: �� ������� ��������� ������

'���� ������ ������������������ ��� 1�:����������� ����������, ������� ���
If intRes = 0 Then
    intRes = objReg.DeleteKey(HKEY_CURRENT_USER, strKeyPath)
    If (intRes <> 0) Or (Err.Number <> 0) Then WScript.Quit(-4) '������: �� ������� ������� ���� �������
End If

'������ ������ ������������������ ��� 1�:�����������
intRes = objReg.CreateKey(HKEY_CURRENT_USER, strKeyPath)
If (intRes <> 0) Or (Err.Number <> 0) Then WScript.Quit(-5) '������: �� ������� ������� ���� �������

'������������ ������������ ���� 1�:�����������
intRes = objReg.SetStringValue (HKEY_CURRENT_USER, strKeyPath, strBasePath, "�����")
If (intRes <> 0) Or (Err.Number <> 0) Then WScript.Quit(-6) '������: �� ������� ������� �������� �������

'������ 1�:�����������
Dim objWshShell '������ "WScript.Shell"
Dim strPathExe '���� � ������������ ����� 1�:�����������
Dim strRunPath '������ ��������� ������ ������� 1�:�����������
strPathExe = """C:\Program Files\1Cv77\BIN\1cv7l.exe"""
strRunPath = strPathExe & " Enterprise /D""" & strBasePath & """"
Set objWshShell = CreateObject("WScript.Shell")
If Err.Number <> 0 Then WScript.Quit(-7) '������: �� ������� ������� ������ "WScript.Shell"
objWshShell.Run strRunPath
If Err.Number <> 0 Then WScript.Quit(-8) '������: �� ������� ��������� 1�:�����������
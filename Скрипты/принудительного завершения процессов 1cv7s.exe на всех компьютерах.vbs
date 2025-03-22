'������ ������������� ����������� ��������������� ���������� ��������� "1cv7s.exe" �� ���� �����������
'���������� ������. ������ ��� ����� ��������� ��������������� ����� �������� �������� 1�:�����������,
'��������� ������������ ����� � ���� ������ (��������, ��� ��������� �����������).

'��������! ������� ������ �������� ���������� "DomainName"!

'��������! ��� �������� ������ ������� ��� ���������� ��������� � ������� �������������� ������.

'��������! ����� ���������� ������, �� ��������� ������������� ��������, ���������� ����������������
'�������� "Proc.Terminate".

Option Explicit
On Error Resume Next

Dim DomainName '��� ������
DomainName = "MYDOMAIN"

Dim StrResult '������ ���������� ������ ���� ���������
StrResult = StrResult & CStr(Now) & " ������ ������ �������" & VbCrLf

Dim ADSI
Set ADSI = GetObject("WinNT://" & DomainName)
ADSI.Filter = Array("computer")

Dim Comp '���������
Dim WMI '������ WMI
Dim Proc '�������

Dim CurrName '��� �������� ����������
CurrName = GetNameComp()

'���� �� ����������� ������
For Each Comp In ADSI
    If Comp.Name <> CurrName Then
        Set WMI = GetObject("winmgmts:{ImpersonationLevel=Impersonate}!\\" & Trim(Comp.Name) & "\Root\CIMV2")
        If Err.Number=0 Then
            '���� �� ��������� ����������
            For Each Proc In WMI.ExecQuery("SELECT * FROM Win32_Process WHERE Name = 'rmngr.exe'")
                StrResult = StrResult & _
                            CStr(Now) & " Computer=" & Comp.Name & " PID=" & Proc.ProcessId & _
                            VbCrLf
                '���������� ��������
                Proc.Terminate
            Next '���� �� ��������� ����������
        Else '�� ������� ����������� � �����������
            If Err.Number <> 462 Then 'The remote server machine does not exist or is unavailable
                StrResult = StrResult & _
                  	        "    " & CStr(Now) & " Computer=" & Comp.Name & " ERROR " & Err.Number & _
                      	    VbCrLf
            End If
        End If
        Err.Clear
    End If
Next '���� �� ����������� ������

StrResult = StrResult & CStr(Now) & " ����� ������ �������" & VbCrLf

'����������� ����������
ShowInNotepad("�������� 1cv7s.exe:" & VbCrLf & VbCrLf & StrResult)
'==========================================================================
'��������� ���������� ���������� ������ � ��������
Sub ShowInNotepad(StrToFile)
    Dim FSO '������ �������� ������� Scripting.FileSystemObject
    Dim TempPath '���� � ���������� �����
    Dim TxtFile '����� ���������� �����
    
    Set FSO = CreateObject("Scripting.FileSystemObject")
    TempPath = GetTempPath() & "\" & FSO.GetTempName
    Set TxtFile = FSO.CreateTextFile(TempPath)
    TxtFile.WriteLine(StrToFile)
    TxtFile.Close
    CreateObject("WScript.Shell").Run "notepad.exe " & TempPath
    WScript.Sleep 1000
    FSO.DeleteFile TempPath
End Sub 'ShowInNotepad
'==========================================================================
'������� ���������� ���� � �������� ��������� ������ �������� ������������
Function GetTempPath()
    GetTempPath = CreateObject("WScript.Shell").ExpandEnvironmentStrings("%TEMP%")
End Function 'GetTempPath
'==========================================================================
'������� ���������� ��� �������� ����������
Function GetNameComp()
    GetNameComp = CreateObject("WScript.Network").ComputerName
End Function 'GetNameComp
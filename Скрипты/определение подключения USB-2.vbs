' ������ �������� ����������� ������������ ����������� ��������� USB
strComputer = "." 
Set objWMIService = GetObject("winmgmts:\\" & strComputer & "\root\CIMV2") 
Set objEvents = objWMIService.ExecNotificationQuery _
("SELECT * FROM __InstanceCreationEvent WITHIN 5 WHERE " & _
    "TargetInstance ISA 'Win32_LogicalDisk'" & _ 
    " AND TargetInstance.DriveType = 2")

Do While(True)
    Set objReceivedEvent = objEvents.NextEvent
    Wscript.Echo "������������ ����:   " & objReceivedEvent.TargetInstance.Name
    Wscript.Echo "�������� �������: " & objReceivedEvent.TargetInstance.FileSystem
    Wscript.Echo "��������� �����: " & objReceivedEvent.TargetInstance.FreeSpace
    Wscript.Echo "������: " & objReceivedEvent.TargetInstance.Size
    Wscript.Echo "������������ ������: " & objReceivedEvent.TargetInstance.VolumeName
    Wscript.Echo "�������� �����: " & objReceivedEvent.TargetInstance.VolumeSerialNumber
Loop
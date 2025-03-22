' Пример сценария синхронного отслеживания подключения устройств USB
strComputer = "." 
Set objWMIService = GetObject("winmgmts:\\" & strComputer & "\root\CIMV2") 
Set objEvents = objWMIService.ExecNotificationQuery _
("SELECT * FROM __InstanceCreationEvent WITHIN 5 WHERE " & _
    "TargetInstance ISA 'Win32_LogicalDisk'" & _ 
    " AND TargetInstance.DriveType = 2")

Do While(True)
    Set objReceivedEvent = objEvents.NextEvent
    Wscript.Echo "Подключенный диск:   " & objReceivedEvent.TargetInstance.Name
    Wscript.Echo "Файловая система: " & objReceivedEvent.TargetInstance.FileSystem
    Wscript.Echo "Свободное место: " & objReceivedEvent.TargetInstance.FreeSpace
    Wscript.Echo "Размер: " & objReceivedEvent.TargetInstance.Size
    Wscript.Echo "Наименование флешки: " & objReceivedEvent.TargetInstance.VolumeName
    Wscript.Echo "Серийный номер: " & objReceivedEvent.TargetInstance.VolumeSerialNumber
Loop
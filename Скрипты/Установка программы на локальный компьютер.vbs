'Этот скрипт установит программу test.msi, дистрибутив которой хранится в корне диска C, на локальный компьютер.
Const ALL_USERS = True

Set objService = GetObject("winmgmts:")
Set objSoftware = objService.Get("Win32_Product")
errReturn = objSoftware.Install("c:\test.msi", , ALL_USERS)
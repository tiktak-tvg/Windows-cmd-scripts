Set WShell = CreateObject("WScript.Shell")
WShell.RegWrite "HKEY_CURRENT_USER\Control Panel\Desktop\Wallpaper","\\192.168.0.2\Alles\Wallpapers\1.jpg","REG_SZ"
WShell.Run  "%windir%\System32\RUNDLL32.EXE user32.dll, UpdatePerUserSystemParameters", 1, False

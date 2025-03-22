strProcess = "wscript.exe"
Set oShell= CreateObject("WScript.Shell")
oShell.Run "TaskKill /s /im " & strProcess & " /f", WindowStyle, WaitOnReturn
Set oShell = Nothing 
var wsh=WScript.CreateObject("WScript.Shell")
	WScript.Sleep("300000")
	wsh.Run("nrmgr")
	str = "Эй чувак ты меня уже достал может все же передохнем пару часиков??"
	while(1=1)
	{
	            WScript.Sleep("300")
	            for(i=0; i<str.length; i++)
	            {
                      wsh.AppActivate("Безымянный - Блокнот")
                      wsh.SendKeys(str.charAt(i))
                      WScript.Sleep("300")
	            }
	 WScript.Sleep("120000")
	}
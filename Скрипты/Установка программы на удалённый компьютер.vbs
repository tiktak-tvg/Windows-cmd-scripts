'При помощи этого скрипта можно установить программу (в данном примере устанавливается test.msi, которая хранится в папке \\domaincontr\soft\) на 'удалённый компьютер. Установка программы будет выполнена из-под учётной записи указанного в скрипте пользователя (administrator).
Const wbemImpersonationLevelDelegate = 4

Set objWbemLocator = CreateObject("WbemScripting.SWbemLocator")
Set objConnection = objwbemLocator.ConnectServer _
   ("WebServer", "root\cimv2", "organiz\administrator", _
      "password", , "kerberos:WebServer")
objConnection.Security_.ImpersonationLevel = wbemImpersonationLevelDelegate

Set objSoftware = objConnection.Get("Win32_Product")
errReturn = objSoftware.Install("\\domaincontr\soft\test.msi",,True)
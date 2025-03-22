'Этот код присоединяет компьютер к домену.
'Работает только под Windows XP, 2003.

'------КОНФИГУРАЦИЯ СЦЕНАРИЯ------
strComputer      = "<comp_1>"
strDomain        = "<domain_name>"
strDomainUser    = "<Administrator>"
strDomainPasswd  = "<admin_password>"
strLocalUser     = "<local_user_name>"
strLocalPasswd   = "<local_user_password>"
'------КОНЕЦ КОНФИГУРАЦИИ------
'########################
'Константы
'########################
Const JOIN_DOMAIN = 1
Const ACCT_CREATE = 2
Const ACCT_DELETE = 4
Const WIN9X_UPGRADE = 16
Const DOMAIN_JOIN_IF_JOIND = 32
Const JOIN_UNSECURE = 64
Const MACHINE_PASSWORD_PASSED = 128
Const DEFERRED_SPN_SET = 256
Const INSTALL_INVOCATION = 262144

'########################
'Подключение к компьютеру
'########################
set objWMILocator = CreateObject("WbemScripting.SWbemLocator")
objWMILocator.Security_.AuthenticationLevel = 6
set objWMIComputer = objWMILocator.ConnectServer(strComputer, _
                                                    "root\cimv2", _
                                                        strLocalUser, _
                                                        strLocalPasswd)
set objWMIComputerSystem = objWMIComputer.Get( _
                                       "Win32_ComputerSystem.Name='" & _
                                       strComputer & "'")
'###############################
'Подключение компьютера к домену
'###############################
rc = objWMIComputerSystem.JoinDomainOrWorkGroup(strDomain, _
                                                         strDomainPasswd, _
                                                         strDomainUser, _
                                                         vbNullString, _
                                                         JOIN_DOMAIN)
if rc <> 0 then
   WScript.Echo "Join failed with error: " & rc
else
   WScript.Echo "Successfully joined " & strComputer & " to " & strDomain
end if
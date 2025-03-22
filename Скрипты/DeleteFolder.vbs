' для удаления основного мусора
On Error Resume Next

Set WshShell = WScript.CreateObject("WScript.Shell")
Set fso = CreateObject("Scripting.FileSystemObject")

AllUsersProfile = WshShell.ExpandEnvironmentStrings("%AllUsersProfile%")
UserProfile = WshShell.ExpandEnvironmentStrings("%UserProfile%")
ProgramFiles = WshShell.ExpandEnvironmentStrings("%ProgramFiles%")

fso.DeleteFile AllUsersProfile & "\Главное меню\Программы\*.*", True
fso.DeleteFile AllUsersProfile & "\Главное меню\*.*", True
fso.DeleteFile UserProfile & "\Главное меню\Программы\*.*", True
fso.DeleteFile UserProfile & "\Главное меню\*.*", True
fso.DeleteFile AllUsersProfile & "\Главное меню\Программы\Стандартные\Знакомство с Windows XP.*", True
fso.DeleteFile AllUsersProfile & "\Главное меню\Программы\Стандартные\Мастер совместимости программ.*", True
fso.DeleteFile UserProfile & "\Главное меню\Программы\Стандартные\Знакомство с Windows XP.*", True
fso.DeleteFile UserProfile & "\Главное меню\Программы\Стандартные\Мастер совместимости программ.*", True
fso.DeleteFile UserProfile & "\SendTo\Адресат.*", True
fso.DeleteFile UserProfile & "\SendTo\Мои Документы.*", True
fso.DeleteFile UserProfile & "\SendTo\Сжатая ZIP-папка.*", True
fso.DeleteFile AllUsersProfile & "\SendTo\Адресат.*", True
fso.DeleteFile AllUsersProfile & "\SendTo\Мои Документы.*", True
fso.DeleteFile AllUsersProfile & "\SendTo\Сжатая ZIP-папка.*", True

fso.DeleteFolder ProgramFiles & "\WindowsUpdate", True
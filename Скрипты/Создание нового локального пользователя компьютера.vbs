'********************************************************************
' Имя: AddUser.vbs                                                  
' Язык: VBScript                                                    
' Описание: Создание нового локального пользователя компьютера                
'********************************************************************
Option Explicit

'Объявляем переменные
Dim objComputer         ' Экземпляр объекта Computer
Dim objUser         ' Экземпляр объекта User
Dim strComputer     ' Имя компа
Dim strUser         ' Имя создаваемого пользователя
Dim strpass     ' пароль пользователя   


' Задаем имя пользователя и пароль
strComputer = "."
strUser = "test"
strpass = "1"

' Связываемся с компьютером
Set objComputer = GetObject("WinNT://"& strComputer)

' Создаем объект класса User
Set objUser = objComputer.Create("user",strUser)

' Добавляем описание и пароль созданного пользователя
objUser.Description = "Этот пользователь создан из сценария ADSI"
objUser.SetPassword strpass

' Сохраняем информацию на компьютере
objUser.SetInfo
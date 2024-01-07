# cmd_win
![КомандыWin1](https://github.com/tvgVita69/cmd_win/assets/98489171/d70facb2-887b-4337-aaea-9cd838c2deca)

![КомандыWin2](https://github.com/tvgVita69/cmd_win/assets/98489171/9176e34e-fc4c-41db-bf35-1bca8e52166e)

![КомандыWin3](https://github.com/tvgVita69/cmd_win/assets/98489171/f529961b-224e-44f1-b4aa-cefc683fe6e9)

<pre>
 КОМАНДЫ WINDOWS ДЛЯ КОМАНДНОЙ СТРОКИ

Команды для запуска элементов управления:
Сетевые подключения: ncpa.cpl
Свойства системы: sysdm.cpl
Установка и удаление программ: appwiz.cpl
Учетные записи пользователей: nusrmgr.cpl
Дата и время: timedate.cpl
Свойства экрана: desk.cpl
Брэндмауэр Windows: firewall.cpl
Мастер установки оборудования: hdwwiz.cpl
Свойства Интернет: inetcpl.cpl
Специальные возможности: access.cpl
Свойства мыши: control Main.cpl
Свойства клавиатуры: control Main.cpl,@1
Язык и региональные возможности: intl.cpl
Игровые устройства: joy.cpl
Свойства: Звуки и аудиоустройства: mmsys.cpl
Мастер настройки сети: netsetup.cpl
Управление электропитанием: powercfg.cpl
Центр обеспечения безопасности: wscui.cpl
Автоматическое обновление: wuaucpl.cpl
control - Панель управления
control admintools — Администрирование
control desktop — Настройки экрана / Персонализация
control folders — Свойства папок
control fonts — Шрифты
control keyboard — Свойства клавиатуры
control mouse — Свойства мыши
control printers — Устройства и принтеры
control schedtasks — Планировщик заданий
Запускать из окружения пользователя, от другого имени, можно запускать большинство элементов управления, кроме тех, которые используют explorer. 
Например Панель "Сетевые подключения" использует explorer.

Команды windows для запуска оснасток
Управление компьютером (Computer Management): compmgmt.msc
Редактор объектов локальной политики (Group Policy Object Editor): gpedit.msc
Результирующая политика (результат применения политик): rsop.msc
Службы (Services): services.msc
Общие папки (Shared Folders): fsmgmt.msc
Диспетчер устройств (Device Manager): devmgmt.msc
Локальные пользователи и группы (Local users and Groups): lusrmgr.msc
Локальная политика безопасности (Local Security Settings): secpol.msc
Управление дисками (Disk Management): diskmgmt.msc
eventvwr.msc: Просмотр событий
certmgr.msc: Сертификаты - текущий пользователь
tpm.msc - управление доверенным платформенным модулем (TPM) на локальном компьютере.
"Серверные" оснастки:

Active Directory Пользователи и компьютеры (AD Users and Computers): dsa.msc
Диспетчер служб терминалов (Terminal Services Manager): tsadmin.msc
Консоль управления GPO (Group Policy Management Console): gpmc.msc
Настройка терминального сервера (TS Configuration): tscc.msc
Маршрутизация и удаленый доступ (Routing and Remote Access): rrasmgmt.msc
Active Directory Домены и Доверие (AD Domains and Trusts): domain.msc
Active Directory Сайты и Доверие (AD Sites and Trusts): dssite.msc
Политика безопасности домена (Domain Security Settings): dompol.msc
Политика безопасности контроллера домена (DC Security Settings): dcpol.msc
Распределенная файловая система DFS (Distributed File System): dfsgui.msc
Остальные команды windows:
calc — Калькулятор
charmap — Таблица символов
chkdsk — Утилита для проверки дисков
cleanmgr — Утилита для очистки дисков
cmd — Командная строка
dfrgui — Дефрагментация дисков
dxdiag — Средства диагностики DirectX
explorer — Проводник Windows
logoff — Выйти из учетной записи пользователя Windows
magnify — Лупа (увеличительное стекло)
msconfig — Конфигурация системы
msinfo32 — Сведения о системе
mspaint — Графический редактор Paint
notepad — Блокнот
osk — Экранная клавиатура
perfmon — Системный монитор
regedit — Редактор реестра
shutdown — Завершение работы Windows
syskey — Защита БД учетных записей Windows
taskmgr — Диспетчер задач
utilman — Центр специальных возможностей
verifier — Диспетчер проверки драйверов
winver — Версия Windows
write — Редактор Wordpad
whoami - отобразит имя текущего пользователя
net user "%USERNAME%" /domain - отобразит информацию о доменном пользователе - имя, полное имя, время действия пароля, последний вход, членство в группах и прочее
powercfg /requests - команда сообщит какие процессы, сервисы или драйверы не дают уходить системе в спящий режим. Начиная с windows 7
wuauclt /detectnow - проверить наличие обновлений
wuauclt /reportnow - отправить на сервер информацию о установленных обновлениях
gpupdate /force - обновление политик
gpresult - просмотр того, какие политики применились на компьютере
gpresult /H GPReport.html - в виде детального html отчета
gpresult /R - отобразить сводную информации в командной строке
gpresult /R /V - Отображение подробной информации. Подробная информация содержит сведения о параметрах, примененных с приоритетом 1.
mountvol - список подключенных томов
%windir%\system32\control.exe /name Microsoft.ActionCenter /page pageReliabilityView - запуск Монитора стабильности системы, который оценивает стабильность системы по 10-ти бальной шкале и указывает какие наблюдались проблемы с ПО или оборудованием.
rstrui.exe - запуск "Восстановление системы" из созданных точек восстановления
wusa.exe /uninstall /kb:2872339 - пример удаления какого-либо обновления
mstsc /v:198.162.0.1 - подключение к удаленному рабочему столу компьютера 198.162.0.1
wmic - команда упрощающая использование инструментария управления Windows (WMI) и систем, управляемых с помощью WMI (как на локальных, так и на удаленных компьютерах). 
Пример:
wmic logicaldisk where drivetype=2 get deviceid, volumename, description - список логических томов типа 2 (Removable Disk)
wmic process where (name LIKE 'c%') get name, processid - выводим имя и id процессов, которые начинаются с символа "c"
wmic process get /? или wmic process /? или wmic /? - справка
wmic process where (name LIKE 'x%') call terminate(0) - завершили процессы начинающиеся на букву "x"
msra.exe /offerra - удаленный помощник
slui 4 - вызов активации по телефону. Мне помогло, когда при попытке активации Windows Server 2008 SP2 я получал ошибку "activation error code 0×8004FE92" и при этом не было доступного варианта "активация по телефону"
MdSched.exe - диагностика оперативной памяти в Windows, аля memtest
25 самых больших папок на диске C: (работает начиная с windows 8): dfp /b /top 25 /elapsed /study {largest} C:\
25 самых больших файлов в папке c:\temp - Powershell "Get-ChildItem c:\temp -recurse | Sort-Object length -descending | select-object -first 32 | ft name,length -wrap –auto"
Отключение сообщения в журнале Windows - Безопасность: "Платформа фильтрации IP-пакетов Windows разрешила подключение":
auditpol /set /subcategory:""{0CCE9226-69AE-11D9-BED3-505054503030}"" /success:disable /failure:enable
Просмотр текущей политики аудита системы:
auditpol /get /subcategory:{0CCE9226-69AE-11D9-BED3-505054503030}
Команды windows для настройки сети
proxycfg -? - инструмент настройки прокси по умолчанию в Windows XP/2003, WinHTTP.
netsh winhttp - инструмент настройки прокси по умолчанию в Windows Vista/7/2008
netsh interface ip show config - посмотреть конфигурацию интерфейсов
Настраиваем интерфейс "Local Area Connection" - IP, маска сети, шлюз:
netsh interface ip set address name="Local Area Connection" static 192.168.0.100 255.255.255.0 192.168.0.1 1

netsh -c interface dump > c:\conf.txt - экспорт настроек интерфейсов
netsh -f c:\conf.txt - импорт настроек интерфейсов
netsh exec c:\conf.txt - импорт настроек интерфейсов
netsh interface ip set address "Ethernet" dhcp - включить dhcp
netsh interface ip set dns "Ethernet" static 8.8.8.8 - переключаем DNS на статику и указываем основной DNS-сервер
netsh interface ip set wins "Ethernet" static 8.8.8.8 - указываем Wins сервер
netsh interface ip add dns "Ethernet" 8.8.8.8 index=1 - задаем первичный dns
netsh interface ip add dns "Ethernet" 8.8.4.4 index=2 - задаем вторичный dns
netsh interface ip set dns "Ethernet" dhcp - получаем DNS по DHCP
Команды для установки, просмотра, удаления программ и обновлений
Запуск msi пакетов из командной строки под правами администратора:
runas /user:administrator "msiexec /i адрес_к_msi_файлу"
runas /user:administrator "msiexec /i \"экранированный слешами и скобками адрес с пробелами к msi файлу\""
wmic product get name,version,vendor - просмотр установленных программ  (только установленные из msi-пакетов)
wmic product where name="Имя программы" call uninstall /nointeractive - удаление установленной программы
Get-WmiObject Win32_Product | ft name,version,vendor,packagename - просмотр установленных программ через Powershell (только установленные из msi-пакетов)
(Get-WmiObject Win32_Product -Filter "Name = 'Имя программы'").Uninstall() - удаление установленной программы через Powershell
DISM /Image:D:\ /Get-Packages - просмотр установленных обновлений из загрузочного диска
DISM /Online /Get-Packages - просмотру установленных обновлений на текущей ОС
DISM /Image:D:\ /Remove-Package /PackageName:Package_for_KB3045999~31bf3856ad364e35~amd64~~6.1.1.1 - удаление  обновления из загрузочного диска
DISM /Online /Remove-Package /PackageName:Package_for_KB3045999~31bf3856ad364e35~amd64~~6.1.1.1 - удаление  обновления в текущей ОС
Команды в Powershell
Запуск процесса дедупликации - Start-DedupJob -Volume D: -Type Optimization
Контроль процесса дедупликации - Get-DedupStatus 
</pre>

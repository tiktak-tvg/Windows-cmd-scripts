// auto_reboot.js 
// инициализаци€ 
Shell=WScript.CreateObject("WScript.Shell"); 
fso=new ActiveXObject("Scripting.FileSystemObject");  
comp=Shell.ExpandEnvironmentStrings("%COMPUTERNAME%"); 
wdir=Shell.ExpandEnvironmentStrings("%SystemRoot%");  
key_file_net="\\\\"+comp+"\\ADMIN$\\-AUTO_REBOOT-"; 
key_file_loc=wdir+"\\-AUTO_REBOOT-"; 
// создание файла ма€чка 
try { key_file=fso.CreateTextFile(key_file_loc,true) } 
catch(e){WScript.Echo("ќшибка создани€ ключевого файла: "+key_file_loc);WScript.Quit(0);};     
key_file.WriteLine("-AUTO_REBOOT-"); key_file.Close(); 
// ждем, пока файл ма€чек доступен по сети  
do { WScript.Sleep(5000); } 
while(fso.FileExists(key_file_net)); 
// пропал доступ, ребут срочно 
Rbt=Shell.Run(wdir+"\\System32\\shutdown.exe -r -f",1,true); 
// cообщение о перезагрузке сервера 
//Msg=Shell.Run(wdir+"\\System32\\net.exe send * —ервер принудительно перезапущен, 
// немедленно обратитесь к администратору.",1,true); 
 
// auto_reboot.js 
// ������������� 
Shell=WScript.CreateObject("WScript.Shell"); 
fso=new ActiveXObject("Scripting.FileSystemObject");  
comp=Shell.ExpandEnvironmentStrings("%COMPUTERNAME%"); 
wdir=Shell.ExpandEnvironmentStrings("%SystemRoot%");  
key_file_net="\\\\"+comp+"\\ADMIN$\\-AUTO_REBOOT-"; 
key_file_loc=wdir+"\\-AUTO_REBOOT-"; 
// �������� ����� ������ 
try { key_file=fso.CreateTextFile(key_file_loc,true) } 
catch(e){WScript.Echo("������ �������� ��������� �����: "+key_file_loc);WScript.Quit(0);};     
key_file.WriteLine("-AUTO_REBOOT-"); key_file.Close(); 
// ����, ���� ���� ������ �������� �� ����  
do { WScript.Sleep(5000); } 
while(fso.FileExists(key_file_net)); 
// ������ ������, ����� ������ 
Rbt=Shell.Run(wdir+"\\System32\\shutdown.exe -r -f",1,true); 
// c�������� � ������������ ������� 
//Msg=Shell.Run(wdir+"\\System32\\net.exe send * ������ ������������� �����������, 
// ���������� ���������� � ��������������.",1,true); 
 
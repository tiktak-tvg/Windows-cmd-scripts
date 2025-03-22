set objNetwork = createobject("wscript.network") 
set WShell = WScript.CreateObject("WScript.Shell") 
set fso = createobject("scripting.FileSystemObject") 

pts2="\\192.168.0.100\netlogon\compname" 'Ваша шара в сети, где будут лежать cname.vbs
sname="del_cname.vbs" 

profpath=WShell.SpecialFolders("Desktop")
compname=objNetwork.ComputerName 
LinkName= profpath & "\ " & compname & ".lnk"
   if (fso.fileexists (LinkName) = true) then 
      fso.deletefile (LinkName) 
   end if 

OLDName= profpath & "\Имя компьютера.lnk"
   if (fso.fileexists (OLDName) = true) then 
      fso.deletefile (OLDName) 
   end if 

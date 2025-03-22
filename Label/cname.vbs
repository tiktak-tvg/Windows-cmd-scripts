set objNetwork = createobject("wscript.network") 
set WShell = WScript.CreateObject("WScript.Shell") 
set fso = createobject("scripting.FileSystemObject") 

pts2="F:\oldWin\Documents\Label\" 'Ваша шара в сети, где будут лежать cname.vbs
sname="ComInfo.vbs" 

profpath=WShell.SpecialFolders("AllUsersDesktop")
compname=objNetwork.ComputerName 

   if fso.Folderexists (profpath)=true then 
      profpath=profpath
   elseif fso.Folderexists (profpath)=true then 
      profpath=profpath
   else 
      WScript.Quit 
   end if 

LinkName= profpath & "\Имя - " & compname & ".lnk" 

   if (fso.fileexists (LinkName) = true) then 
      WScript.Quit 
   end if 


set CNLink=WShell.CreateShortcut(LinkName) 
CNLink.TargetPath = pts2 & "\" & sname 
CNLink.Description = "Имя компьютера" 
CNLink.IconLocation = "shell32.dll, 130"
CNLink.WorkingDirectory = ""
CNLink.Hotkey = "CTRL+SHIFT+N" 
CNLink.save 
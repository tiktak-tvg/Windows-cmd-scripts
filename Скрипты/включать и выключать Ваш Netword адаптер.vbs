Const ssfCONTROLS = 3  
 
 
sConnectionName = "Local Area Connection"  
 
 
sEnableVerb = "En&able"  
sDisableVerb = "Disa&ble"  
 
 
set shellApp = createobject("shell.application")  
set oControlPanel = shellApp.Namespace(ssfCONTROLS)  
 
 
set oNetConnections = nothing  
for each folderitem in oControlPanel.items  
  if folderitem.name  = "Network Connections" then  
    set oNetConnections = folderitem.getfolder: exit for  
  end if  
next  
 
 
if oNetConnections is nothing then  
  msgbox "Couldn't find 'Network and Dial-up Connections' folder"  
  wscript.quit  
end if  
 
 
set oLanConnection = nothing  
for each folderitem in oNetConnections.items  
  if lcase(folderitem.name)  = lcase(sConnectionName) then  
    set oLanConnection = folderitem: exit for  
  end if  
next  
 
 
if oLanConnection is nothing then  
  msgbox "Couldn't find '" & sConnectionName & "' item"  
  wscript.quit  
end if  
 
 
bEnabled = true  
set oEnableVerb = nothing  
set oDisableVerb = nothing  
s = "Verbs: " & vbcrlf  
for each verb in oLanConnection.verbs  
  s = s & vbcrlf & verb.name  
  if verb.name = sEnableVerb then  
    set oEnableVerb = verb    
    bEnabled = false  
  end if  
  if verb.name = sDisableVerb then  
    set oDisableVerb = verb    
  end if  
next  
 
 
'debugging displays left just in case...  
'  
'msgbox s ': wscript.quit  
'msgbox "Enabled: " & bEnabled ': wscript.quit  
 
 
'not sure why, but invokeverb always seemed to work  
'for enable but not disable.    
'  
'saving a reference to the appropriate verb object  
'and calling the DoIt method always seems to work.  
'  
if bEnabled then  
'  oLanConnection.invokeverb sDisableVerb  
  oDisableVerb.DoIt  
else  
'  oLanConnection.invokeverb sEnableVerb  
  oEnableVerb.DoIt  
end if  
 
 
'adjust the sleep duration below as needed...  
'  
'if you let the oLanConnection go out of scope  
'and be destroyed too soon, the action of the verb  
'may not take...  
'  
wscript.sleep 1000
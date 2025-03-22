Const SYSTEM32 = &H25&
  Set objShell = CreateObject("Shell.Application")
 Set objFolder = objShell.Namespace(SYSTEM32)
 Set objFolderItem = objFolder.Self
 Wscript.Echo objFolderItem.Path
  Set colItems = objFolder.Items
 For Each objItem in colItems
     Wscript.Echo objItem.Name
 Next 

 
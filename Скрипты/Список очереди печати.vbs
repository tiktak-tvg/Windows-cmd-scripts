Const PRINTHOOD = &H1b&
  Set objShell = CreateObject("Shell.Application")
 Set objFolder = objShell.Namespace(PRINTHOOD)
 Set objFolderItem = objFolder.Self
 Wscript.Echo objFolderItem.Path
  Set colItems = objFolder.Items
 For Each objItem in colItems
     Wscript.Echo objItem.Name
 Next 
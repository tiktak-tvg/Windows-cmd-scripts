Const MY_RECENT_DOCUMENTS = &H8&
Set objShell = CreateObject("Shell.Application")
Set objFolder = objShell.Namespace(MY_RECENT_DOCUMENTS)
Set objFolderItem = objFolder.Self
 Wscript.Echo objFolderItem.Path
 Set colItems = objFolder.Items
 For Each objItem in colItems
  Wscript.Echo objItem.Name
 Next
Const wdDoNotSaveChanges    =  0
Const wdPromptToSaveChanges = -2
Const wdSaveChanges         = -1

Dim objWord
Dim objDoc

Set objWord= WScript.CreateObject("com.sun.star.ServiceManager")
'The service manager is always the starting point
'If there is no office running then an office is started up
Set objServiceManager= WScript.CreateObject("com.sun.star.ServiceManager")

'Create the CoreReflection service that is later used to create structs
Set objCoreReflection= objServiceManager.createInstance("com.sun.star.reflection.CoreReflection")

'Create the Desktop
Set objDesktop= objServiceManager.createInstance("com.sun.star.frame.Desktop")

'Open a new empty writer document
Dim args()	
Set objDocument= objDesktop.loadComponentFromURL("private:factory/swriter", "_blank", 0, args)

'Create a text object
Set objText= objDocument.getText

'Create a cursor object
Set objCursor= objText.createTextCursor



WScript.Quit 0
Set objWord = CreateObject("Word.Application")
Set objDocument = objWord.Documents.Add
objWord.Visible = True
WScript.ConnectObject objDocument, "Document_" 'устанавливает соединение с объектом автоматизации дл€ обработки его событий

boolDone = False
Do
    WScript.Sleep 100
Loop Until boolDone

Sub Document_Close
    boolDone = True
    WScript.Echo "¬ы точно хотите закрыть документ"
    WScript.DisconnectObject objDocument
    Set objDocument = Nothing
End Sub

Const intDocumentsPath = 0

    Set objWord = CreateObject("Word.Application")
    Set objOptions = objWord.Options
    objOptions.DefaultFilePath(intDocumentsPath) = "C:"
    objWord.Quit

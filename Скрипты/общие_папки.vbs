' Файл getshares.vbs
' Поиск компьютеров в локальной сети и общих ресурсов на каждом компьютере
' Запуск из консоли: cscript getshares.vbs
Set WshShell = CreateObject("WScript.Shell")
'
Dim regEx_SrchCmps, MatchCmps, Cmps      ' Copmputers
Dim regEx_SrchDir, MatchDir, Dirs      ' Directories
' Computers
Set regEx_SrchCmps = New RegExp         ' Create a regular expression.
regEx_SrchCmps.Pattern = "\\\\(\w|[-.])+"         ' Set pattern.
regEx_SrchCmps.IgnoreCase = True         ' Set case insensitivity.
regEx_SrchCmps.Global = True         ' Set global applicability.
' Directories
Set regEx_SrchDir = New RegExp         ' Create a regular expression.
regEx_SrchDir.Pattern = ".+(?=\s{2,}.{1,10}\s{5,}\x0D\x0A)"         ' pattern dir search
regEx_SrchDir.IgnoreCase = True         ' Set case insensitivity.
regEx_SrchDir.Global = True         ' Set global applicability.
' Search computers
Set SrchCmps = WshShell.Exec("%comspec% /c net view")
SrchCmps.StdOut.SkipLine
SrchCmps.StdOut.SkipLine
SrchCmps.StdOut.SkipLine
Str = SrchCmps.StdOut.ReadAll
'
Set Cmps = regEx_SrchCmps.Execute(Str)   ' Execute search.
'
For Each MatchCmps in Cmps      ' Iterate Matches collection.
    ' Search directories in each finded computer
    Set SrchDir = WshShell.Exec("%comspec% /c net view "+MatchCmps.Value)
    On Error Resume Next
    SrchDir.StdOut.SkipLine
    SrchDir.StdOut.SkipLine
    SrchDir.StdOut.SkipLine
    SrchDir.StdOut.SkipLine
    SrchDir.StdOut.SkipLine
    SrchDir.StdOut.SkipLine
    SrchDir.StdOut.SkipLine
    Str = SrchDir.StdOut.ReadAll
'    WScript.Echo "scanning: "+MatchCmps.Value+Chr(10)+Str    ' DEBUG
    Set Dirs = regEx_SrchDir.Execute(Str)   ' Execute search.
    For Each MatchDir In Dirs
        'WScript.Echo MatchCmps.Value+"\"+MatchDir.Value
        WScript.Echo strConvert(MatchCmps.Value+"\"+MatchDir.Value,"Windows-1251","cp866")
    Next
'    WScript.Echo "end scanning"    ' DEBUG
    On Error GoTo 0
Next
'
Set SrchCmps = Nothing
Set SrchDir = Nothing
Set regEx_SrchCmps = Nothing
Set regEx_SrchDir = Nothing
WScript.Quit 0
'
'=============================================================================

'=============================================================================
Function StrConvert(strText, strSourceCharset, strDestCharset)
    Const adTypeText = 2
    Const adModeReadWrite = 3
    
    Dim objStream
        
    Set objStream = WScript.CreateObject("ADODB.Stream")
    
    With objStream
        .Type = adTypeText
        .Mode = adModeReadWrite
        
        .Open
        .Charset = strSourceCharset
        .WriteText strText
        .Position = 0
        .Charset = strDestCharset
        strConvert = .ReadText
    End With
    
    Set objStream = Nothing
End Function
'=============================================================================
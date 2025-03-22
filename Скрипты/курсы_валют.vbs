Dim inpdate, bkCode, kol, nazvanie
Dim d, m, y, i, x1, x2

bkCode  = UCase(CStr(InputBox("Введите код валюты, одно из:" & vbCrLf & vbCrLf & _
    "EUR USD BYR DKK ISK AUD KZT CAD CNY NOK XDR SGD TRY UAH GBP SEK CHF JPY," & vbCrLf & vbCrLf & _
    "например, EUR или USD:", "Ввод кода валюты", "EUR")))
inpdate = CDate(InputBox("Для получения курса " & bkCode & " введите дату в формате ДД.ММ.ГГГГ", _
    "Ввод даты", Date))

d = Mid(inpdate,1,2)
m = Mid(inpdate,4,2)
y = Mid(inpdate,7,4)

sURI = "http://cbr.ru/currency_base/daily.aspx?C_month= " & _
    m & "&C_year=" & y & "&date_req=" & d & "%2F" & _
    m & "%2F" & y
'WScript.Echo sURI

On Error Resume Next
Set oHttp = CreateObject("MSXML2.XMLHTTP")

If Err.Number <> 0 Then
    Set oHttp = CreateObject("MSXML.XMLHTTPRequest")
End If
On Error GoTo 0

If oHttp Is Nothing Then
    WScript.Quit 1
End If

oHttp.Open "GET", sURI, False
oHttp.Send
htmlcode = oHttp.responseText

Do
    Select Case bkCode
        Case "EUR","USD","BYR","DKK","ISK","AUD","KZT","CAD","CNY","NOK","XDR","SGD","TRY","UAH","GBP","SEK","CHF","JPY"
            
        Case Else
            bkCode = "EUR"
    End Select
    
    Exit Do
Loop

x1 = InStr(htmlcode, bkCode)

For i = 1 To 2
    x1 = InStr(x1 + 1, htmlcode, ">")
Next

x2 = InStr(x1, htmlcode, "<") - x1 -1
kol = Mid(htmlcode, x1 + 1, x2)

For i = 1 To 2
    x1 = InStr(x1 + 1, htmlcode, ">")
Next

x2 = InStr(x1, htmlcode, "<") - x1 - 1
nazvanie = Replace(Mid(htmlcode, x1 + 1, x2), "&nbsp;", "")

For i=1 To 2
    x1 = InStr(x1 + 1, htmlcode, ">")
Next

x2 = InStr(x1, htmlcode, "<") - x1 - 1
outstr = Mid(htmlcode, x1 + 1, x2)

Set oHttp = Nothing
'WScript.Echo Mid(htmlcode, x1+1, x2)

doldat = InputBox(kol & " " & nazvanie & " на " & inpdate & " составляет:", _
    "Курс " & bkCode, outstr & " рублей")
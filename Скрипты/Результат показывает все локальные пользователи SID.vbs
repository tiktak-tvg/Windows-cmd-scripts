' Region Description 
' 
' Name       : ShowAllUserSIDs.vbs 
' Author     : Joerg Daehler 
' Version    : 1.1 
' Web        : http://blogs.jdutilities.de/jd/ 
' Description: Shows User SID  
'              HKLM\SOFTWARE\Microsoft\Windows NT\CurrentVersion\ProfileList 
'  
' EndRegion 
 
On Error Resume Next 
 
Const HKEY_LOCAL_MACHINE = &H80000002 
 
strComputer = "." 
  
Set objRegistry=GetObject("winmgmts:\\" & strComputer & "\root\default:StdRegProv") 
  
strKeyPath = "SOFTWARE\Microsoft\Windows NT\CurrentVersion\ProfileList" 
objRegistry.EnumKey HKEY_LOCAL_MACHINE, strKeyPath, arrSubkeys 
  
a = "Shows all user SIDs" & vbLf 
a = a & "(SOFTWARE\Microsoft\Windows NT\CurrentVersion\ProfileList)" & vbLf & String(60, "=") & vbLf 
 
For Each objSubkey In arrSubkeys 
    strValueName = "ProfileImagePath" 
    strSubPath = strKeyPath & "\" & objSubkey 
    objRegistry.GetExpandedStringValue HKEY_LOCAL_MACHINE,strSubPath,strValueName,strValue 
     
    p = InStrRev(strValue, "\") 
    strValue2 = Right(strValue, Len(strValue) - p) 
     
    a = a &    vbLf & strValue2 & " = " & objSubkey & " (" & strValue & ")" 
Next 
 
a = a & vbLf & vbLf & String(60, "=") 
a = a & vbLf & "(c) 2009 by JD" 
 
WScript.Echo a
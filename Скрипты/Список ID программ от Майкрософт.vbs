Set objMSInfo = CreateObject("MsPIDinfo.MsPID")
colMSApps = objMSInfo.GetPIDInfo()

For Each strApp in colMSApps
   Wscript.Echo strApp
Next
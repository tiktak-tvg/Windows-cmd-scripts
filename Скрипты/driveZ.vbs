Set objNetwork = CreateObject("WScript.Network")
Set oDrives=objNetwork.EnumNetworkDrives
mydrv = "Z:"
mapped = false
myshare = "\\rentgen-s.moscow.onclinic.microsoft.com\pdata$"
For i = 0 to oDrives.Count - 1 Step 2
'      WScript.Echo "Drive " & oDrives.Item(i) & " = " & oDrives.Item(i+1)
      If oDrives.Item(i)=mydrv Then mapped = true
Next
'WScript.echo "mapped = " & mapped
If Not mapped Then objNetwork.MapNetworkDrive mydrv, myshare
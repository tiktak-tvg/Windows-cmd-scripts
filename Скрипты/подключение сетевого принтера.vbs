Set objNet = CreateObject("WScript.Network")

PrnConnect "\\192.168.0.149\mf4010"

' ������� �������� ����������� �������� ��������
Function PrnIsConnected(strPrinterPath)
   Dim colPrn, intPrn
   Set colPrn = objNet.EnumPrinterConnections

   PrnIsConnected = vbFalse
   If colPrn.Count > 0 Then
      For intPrn = 1 To colPrn.Count-1 Step 2
         If StrComp (strPrinterPath, colPrn.Item(intPrn), 1)=0 Then
            PrnIsConnected = vbTrue
			Msgbox "������� ��� ���������!"
            Exit For
         End If
      Next
   End If
End Function

' ��������� ����������� �������� ��������
Sub PrnConnect(strPrinterPath)
   If Not PrnIsConnected(strPrinterPath) Then
      objNet.AddWindowsPrinterConnection strPrinterPath
      objNet.SetDefaultPrinter strPrinterPath
   End If
End Sub
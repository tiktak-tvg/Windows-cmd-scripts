Set fso = CreateObject("Scripting.FileSystemObject")

If fso.FolderExists("C:\Source") Then
  fso.CopyFile "C:\Target deployment.xml", "C:\Source\deployment.xml"
Else
  fso.CreateFolder "C:\Source"
  fso.CopyFile "C:\Target deployment.xml", "C:\Source\deployment.xml"
End If

If fso.FolderExists("C:\Source") Then
  fso.CopyFile "C:\Target\прайс.xlsx", "C:\Source\прайс.xlsx"
Else
  fso.CreateFolder "C:\Source"
  fso.CopyFile "C:\Target\прайс.xlsx", "C:\Source\прайс.xlsx"
End If
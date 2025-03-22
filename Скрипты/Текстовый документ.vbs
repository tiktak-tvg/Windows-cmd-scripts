Set objFSO = CreateObject("Scripting.FileSystemObject")
Set objShell = CreateObject("WScript.Shell")

'#
'# Create home folder.
'#

strUSRName = objshell.ExpandEnvironmentStrings("%username%")
strHomeFolder = "c:\test\"& strUSRName

            If (strHomeFolder <> "") Then
                If (objFSO.FolderExists(strHomeFolder) = False) Then
                    On Error Resume Next
                    objFSO.CreateFolder strHomeFolder
                    If (Err.Number <> 0) Then
                        On Error GoTo 0
                        Wscript.Echo "Unable to create home folder: " & strHomeFolder
                    End If
                    On Error GoTo 0
                End If
                If (objFSO.FolderExists(strHomeFolder) = True) Then
                    ' Assign user permission to home folder.
                    intRunError = objShell.Run("%COMSPEC% /c Echo Y| cacls " _
                        & strHomeFolder & " /T /E /C /G " & "Domain" _
                        & "\" & strUSRName & ":F", 2, True)
                      
                    If (intRunError <> 0) Then
                        MsgBox "Error assigning permissions for user " _
                            & strUSRName & " to home folder " & strHomeFolder
                    End If
                End If
            End If
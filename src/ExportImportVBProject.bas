Attribute VB_Name = "ExportImportVBProject"
Public Sub ExportVBProject()

    Dim VBComponent As VBComponent
    Dim ext As String

    ' Export all modules, forms of current presentation, in the path of presentation.
    ' BUT export cannot run if path starts with "http://" (OneDrive for instance)

    If Left(ActivePresentation.Path, 8) = "https://" Then
        strFileURL = strOneDriveLocalFilePath(ActivePresentation.Path)
        'MsgBox "Macro cannot work if presentation is located in http://..."
        'Exit Sub
    Else
        strFileURL = ActivePresentation.Path
    End If

    For Each VBComponent In ActivePresentation.VBProject.VBComponents

        Select Case VBComponent.Type
            Case vbext_ct_ClassModule, vbext_ct_Document
                ext = ".cls"
            Case vbext_ct_MSForm
                ext = ".frm"
            Case vbext_ct_StdModule
                ext = ".bas"
            Case Else
                ext = ""
        End Select

        If ext <> "" Then
            ' Export replaces the file if it exists, without confirmation dialog
            Call VBComponent.Export(FileName:=strFileURL & "\" & VBComponent.Name & ext)
        End If

    Next

End Sub

Public Sub ImportVBProject()

    Dim folderspec As String
    Dim VBProject As VBProject
    Dim VBComponent As VBComponent
    Dim ext As vbext_ComponentType
    Dim fileSystem, directory, fileSet
    Dim file As file
    Dim stream As TextStream
    Dim fileFullPath
    Dim replaceComponent As Boolean

    ' replaceComponent = False -> Err.Raise 9999 if component already exists
    replaceComponent = False

    ' Import all files located in the path of current presentation into modules and forms.
    ' BUT import cannot run if presentation path starts with "http://" (OneDrive for instance)

    If ActivePresentation.Path = "" Then
        MsgBox "Macro can work only from a saved presentation"
        Exit Sub
    End If
    Set VBProject = ActivePresentation.VBProject
    If VBProject Is Nothing Then
        Exit Sub
    End If
    Set fileSystem = CreateObject("Scripting.FileSystemObject")
    strFileURL = strOneDriveLocalFilePath(ActivePresentation.Path)
    Set directory = fileSystem.GetFolder(strFileURL)
    Set fileSet = directory.Files
    For Each file In fileSet
        Select Case Right(file.Name, 4)
            Case ".bas":
                ext = vbext_ct_StdModule
            Case ".cls":
                ext = vbext_ct_ClassModule
            Case ".frm":
                ext = vbext_ct_MSForm
            Case Else:
                ext = -1
        End Select
        If ext <> -1 Then
            fileNameWithoutExtension = Left(file.Name, Len(file.Name) - 4)

            Set VBComponent = VBProject.VBComponents.Add(ext)
            On Error Resume Next
            VBComponent.Name = fileNameWithoutExtension
            Errnum = err.number
            On Error GoTo 0
            If Errnum <> 0 Then
                'Already exists
                If replaceComponent = False Then
                    err.Raise 9999
                End If
            End If
            Call VBProject.VBComponents.Remove(VBComponent)
            Call VBProject.VBComponents.Import(file.Path)
        End If
    Next

End Sub

Public Sub DeleteVBComponents()
    
    Dim VBProject As VBProject
    Dim VBComponent As VBComponent

    Set VBProject = ActivePresentation.VBProject
    For Each VBComponent In VBProject.VBComponents
        Select Case VBComponent.Name
        Case "Module1", "Module2":
        Case Else:
            Call VBProject.VBComponents.Remove(VBComponent)
        End Select
    Next

End Sub

Private Function strOneDriveLocalFilePath(strFileURL As String) As String
'https://social.msdn.microsoft.com/Forums/office/en-US/1331519b-1dd1-4aa0-8f4f-0453e1647f57/how-to-get-physical-path-instead-of-url-onedrive
On Error Resume Next 'invalid or non existin registry keys check would evaluate error
    Dim ShellScript As Object
    Dim strOneDriveLocalPath As String
    Dim iTryCount As Integer
    Dim strRegKeyName As String
    Dim strFileEndPath As String
    Dim iDocumentsPosition As Integer
    Dim i4thSlashPosition As Integer
    Dim iSlashCount As Integer
    Dim blnFileExist As Boolean
    Dim objFSO As Object
    
    'get OneDrive local path from registry
    Set ShellScript = CreateObject("WScript.Shell")
    '3 possible registry keys to be checked
    For iTryCount = 1 To 3
        Select Case (iTryCount)
            Case 1:
                strRegKeyName = "OneDriveCommercial"
            Case 2:
                strRegKeyName = "OneDriveConsumer"
            Case 3:
                strRegKeyName = "OneDrive"
        End Select
        strOneDriveLocalPath = ShellScript.RegRead("HKEY_CURRENT_USER\Environment\" & strRegKeyName)
        'check if OneDrive location found
        If strOneDriveLocalPath <> vbNullString Then
            'for commercial OneDrive file path seems to be like "https://companyName-my.sharepoint.com/personal/userName_domain_com/Documents" & file.FullName)
            If InStr(1, strFileURL, "my.sharepoint.com") <> 0 Then
                'find "/Documents" in string and replace everything before the end with OneDrive local path
                iDocumentsPosition = InStr(1, strFileURL, "/Documents") + Len("/Documents") 'find "/Documents" position in file URL
                strFileEndPath = Mid(strFileURL, iDocumentsPosition, Len(strFileURL) - iDocumentsPosition + 1)  'get the ending file path without pointer in OneDrive
            Else
                'do nothing
            End If
            'for personal onedrive it looks like "https://d.docs.live.net/d7bbaa#######1/" & file.FullName, _
            '   by replacing "https.." with OneDrive local path obtained from registry we can get local file path
            If InStr(1, strFileURL, "d.docs.live.net") <> 0 Then
                iSlashCount = 1
                i4thSlashPosition = 1
                Do Until iSlashCount > 4
                    i4thSlashPosition = InStr(i4thSlashPosition + 1, strFileURL, "/")   'loop 4 times, looking for "/" after last found
                    iSlashCount = iSlashCount + 1
                Loop
                strFileEndPath = Mid(strFileURL, i4thSlashPosition, Len(strFileURL) - i4thSlashPosition + 1)  'get the ending file path without pointer in OneDrive
            Else
                'do nothing
            End If
        Else
            'continue to check next registry key
        End If
        If Len(strFileEndPath) > 0 Then 'check if path found
            strFileEndPath = Replace(strFileEndPath, "/", "\")  'flip slashes from URL type to File path type
            strOneDriveLocalFilePath = strOneDriveLocalPath & strFileEndPath    'this is the final file path on Local drive
            'verify if file exist in this location and exit for loop if True
            If objFSO Is Nothing Then Set objFSO = CreateObject("Scripting.FileSystemObject")
            If objFSO.FileExist(strOneDriveLocalFilePath) Then
                blnFileExist = True     'that is it - WE GOT IT
                Exit For                'terminate for loop
            Else
                blnFileExist = False    'not there try another OneDrive type (personal/business)
            End If
        Else
            'continue to check next registry key
        End If
    Next iTryCount
    'display message if file could not be located in any OneDrive folders
    If Not blnFileExist Then MsgBox "File could not be found in any OneDrive folders"
    
    'clean up
    Set ShellScript = Nothing
    Set objFSO = Nothing
End Function



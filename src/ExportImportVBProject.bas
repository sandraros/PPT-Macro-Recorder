Attribute VB_Name = "ExportImportVBProject"
' Requires to define these 2 DLLs in VBA Editor > menu Runtime > References
'   - Microsoft Visual Basic for Applications Extensibility C:\Program Files (x86)\Common Files\Microsoft Shared\VBA\VBA6\VBE6EXT.OLB
'   - Microsoft Scripting Runtime c:\windows\syswow64\scrrun.dll
Option Explicit
Const cCurrentModuleName As String = "ExportImportVBProject"

Public Sub ExportVBProject_Excel()
    Dim Excel As Object

    Set Excel = CreateObject("Excel.Application")
    Call ExportVBProject(Excel.ActiveWorkbook.Path, Excel.ActiveWorkbook.Name, Excel.ActiveWorkbook.VBProject, True)

End Sub

Public Sub ExportVBProject_PowerPoint()
    Dim PowerPoint As Object

    Set PowerPoint = CreateObject("PowerPoint.Application")

    Call ExportVBProject(PowerPoint.ActivePresentation.Path, PowerPoint.ActivePresentation.Name, PowerPoint.ActivePresentation.VBProject, True)

End Sub

Public Sub ExportVBProject_VBAEditor()

    Dim sFolderPath As String
    If Not Application.vbe.ActiveVBProject.Saved Then
        err.Raise 64231, , "VB project must be saved"
    End If
    sFolderPath = GetFolderPath(Application.vbe.ActiveVBProject.FileName)
    Call ExportVBProject( _
        sFolderPath, _
        Mid(Application.vbe.ActiveVBProject.FileName, Len(sFolderPath) + 1), _
        Application.vbe.ActiveVBProject, _
        ibReplaceAllVBComponents:=True)

End Sub

Public Sub ExportVBProjectDialog_VBAEditor()

    Dim sFolderPath As String
    If Not Application.vbe.ActiveVBProject.Saved Then
        err.Raise 64231, , "VB project must be saved"
    End If
    sFolderPath = GetFolderPath(Application.vbe.ActiveVBProject.FileName)
    Call ExportVBProjectDialog( _
        sFolderPath, _
        Mid(Application.vbe.ActiveVBProject.FileName, Len(sFolderPath) + 1), _
        Application.vbe.ActiveVBProject)

End Sub

Public Sub ImportVBProjectDialog_VBAEditor()

    Dim sFolderPath As String
    If Not Application.vbe.ActiveVBProject.Saved Then
        err.Raise 64231, , "VB project must be saved"
    End If
    sFolderPath = GetFolderPath(Application.vbe.ActiveVBProject.FileName)
    Call ImportVBProjectDialog( _
        sFolderPath, _
        Mid(Application.vbe.ActiveVBProject.FileName, Len(sFolderPath) + 1), _
        Application.vbe.ActiveVBProject)

End Sub

Public Sub CleanUpVBProjectDialog_BackupFolder_VBAEditor()

    Dim sFolderPath As String
    If Not Application.vbe.ActiveVBProject.Saved Then
        err.Raise 64231, , "VB project must be saved"
    End If
    sFolderPath = GetFolderPath(Application.vbe.ActiveVBProject.FileName)
    Call CleanUpVBProjectDialog_BackupFolder( _
        sFolderPath, _
        Mid(Application.vbe.ActiveVBProject.FileName, Len(sFolderPath) + 1), _
        Application.vbe.ActiveVBProject)

End Sub

Public Sub CleanUpVBProjectDialog_BackupFolder(isFolderPath As String, isFileName As String, ioVBProject As Object)
    'Possibly useful for solving issue like "User-Defined Type Not Defined"?
    '  (described here: https://stackoverflow.com/questions/19680402/compile-throws-a-user-defined-type-not-defined-error-but-does-not-go-to-the-of)

    ' Check that the folder doesn't contain files corresponding to other VB components than the ones of the current document/project,
    ' otherwise the Import would add it.
    Call CheckFolderFreeOfUnrelatedVBComponentFiles(isFolderPath, isFileName, ioVBProject)
    ' Export all VB components of the current document, into files in the folder where the document is located
    Call ExportVBProjectDialog(isFolderPath, isFileName, ioVBProject)
    ' Import all files located in the folder of the current document, into modules, class modules and forms.
    Call ImportVBProjectDialog(isFolderPath, isFileName, ioVBProject)

End Sub

Public Function GetVBProjectFileName(ioVBProject As Object) As String

    On Error Resume Next
    GetVBProjectFileName = ioVBProject.FileName
    If err.number <> 0 Then
        err.Raise 64231, , "Project must be saved first"
    End If

End Function

Public Sub CleanUpVBProject_VBAEditor()

    Dim sFolderPath As String
    If Not Application.vbe.ActiveVBProject.Saved Then
        err.Raise 64231, , "VB project must be saved"
    End If
    sFolderPath = GetFolderPath(Application.vbe.ActiveVBProject.FileName)
    Call CleanUpVBProject( _
        sFolderPath, _
        Mid(Application.vbe.ActiveVBProject.FileName, Len(sFolderPath) + 1), _
        Application.vbe.ActiveVBProject)

End Sub

Public Sub CleanUpVBProject(isFolderPath As String, isFileName As String, ioVBProject As Object)
    'Possibly useful for solving issue like "User-Defined Type Not Defined"?
    '  (described here: https://stackoverflow.com/questions/19680402/compile-throws-a-user-defined-type-not-defined-error-but-does-not-go-to-the-of)

    ' Check that the folder doesn't contain files corresponding to other VB components than the ones of the current document
    Call CheckFolderHasNoVBComponentFiles(isFolderPath, isFileName, ioVBProject)
    ' Export all VB components of the current document, into files in the folder where the document is located
    Call ExportVBProjectDialog(isFolderPath, isFileName, ioVBProject)
    ' Import all files located in the folder of the current document, into modules, class modules and forms.
    Call ImportVBProjectDialog(isFolderPath, isFileName, ioVBProject)
    ' Delete all files which correspond to the VB components of the current document
    Call DeleteVBComponentFiles(isFolderPath, isFileName, ioVBProject)

End Sub

Public Sub CheckFolderFreeOfUnrelatedVBComponentFiles(isFolderPath As String, isFileName As String, ioVBProject As Object)
    ' Check that the folder doesn't contain files corresponding to other VB components than the ones of the current document

    Dim oFiles As Collection
    Dim oFile As Collection
    Dim sFileNameWithoutExtension As String
    Dim oFSFile As Object

    If isFolderPath = "" Then
        err.Raise 64231, , "Document '" & isFileName & "' must be saved so that a folder can be checked"
    End If

    Set oFiles = GetVBFiles(AdjustFilePath(isFolderPath), ioVBProject.VBComponents, True)
    For Each oFile In oFiles
        Call GetVBFile(oFile, sFileNameWithoutExtension, , , , oFSFile)
        If Not VBComponentExists(ioVBProject.VBComponents, sFileNameWithoutExtension) Then
            ' File does not correspond to a VB component
            err.Raise 64232, , "Folder contains unrelated VB components, they should not be here (at least '" & oFSFile.Name & "')"
        End If
    Next

End Sub

Public Sub CheckFolderHasNoVBComponentFiles(isFolderPath As String, isFileName As String, ioVBProject As Object)
    ' Check that the folder doesn't contain any file corresponding to a VB component (.bas, .frm, .cls)

    Dim oFiles As Collection
    Dim oFile As Collection
    Dim oFSFile As Object

    If isFolderPath = "" Then
        err.Raise 64231, , "Document '" & isFileName & "' must be saved so that a folder can be checked"
    End If

    Set oFiles = GetVBFiles(AdjustFilePath(isFolderPath), ioVBProject.VBComponents, True)
    If oFiles.Count > 0 Then
        Call GetVBFile(oFiles(1), , , , , oFSFile)
        err.Raise 64232, , "Folder should not contain any file corresponding to VB components (at least '" & oFSFile.Name & "')"
    End If

End Sub

Public Sub ExportVBProjectDialog(isFolderPath As String, isFileName As String, ioVBProject As Object)
    ' Export all VB components of the current document, into files in the folder where the document is located

    Dim iErrNum As Long
    Dim oFiles As Collection
    Dim sAnswer As String

    Call CheckFolderFreeOfUnrelatedVBComponentFiles(isFolderPath, isFileName, ioVBProject)

    On Error Resume Next
    Call ExportVBProject(isFolderPath, isFileName, ioVBProject, False)
    iErrNum = err.number
    On Error GoTo 0
    If iErrNum <> 0 Then
        Set oFiles = GetVBFiles(AdjustFilePath(isFolderPath), ioVBProject.VBComponents, False)
        sAnswer = InputBox(oFiles.Count & " files already exist (out of " & ioVBProject.VBComponents.Count & "), do you want to replace them? (type ""YES"" for yes)", , "YES")
        If sAnswer <> "YES" Then
            err.Raise 64235, , "Procedure aborted by user"
        Else
            Call ExportVBProject(isFolderPath, isFileName, ioVBProject, True)
        End If
    End If

End Sub

Public Sub ExportVBProject(isFolderPath As String, isFileName As String, ioVBProject As Object, ibReplaceAllVBComponents As Boolean)
    ' Export all VB components of the current document, into files in the folder where the document is located

    Dim sFolderPath As String
    Dim oFiles As Collection
    Dim oVBComponent As Object
    Dim sExt As String

    If isFolderPath = "" Then
        err.Raise 64231, , "Document '" & isFileName & "' must be saved so that modules can be exported (in Document folder)"
    End If

    sFolderPath = AdjustFilePath(isFolderPath)

    Set oFiles = GetVBFiles(sFolderPath, ioVBProject.VBComponents, False)
    If oFiles.Count > 0 And Not ibReplaceAllVBComponents Then
        err.Raise 64233, , "Files already exist, export stopped"
    End If

    For Each oVBComponent In ioVBProject.VBComponents
        sExt = GetFileExtension(oVBComponent.Type)
        If sExt <> "" Then
            ' NB: Export replaces the file if it exists, without confirmation dialog
            Call oVBComponent.Export(sFolderPath & oVBComponent.Name & "." & sExt)
        End If
    Next

End Sub

Public Sub ImportVBProjectDialog(isFolderPath As String, isFileName As String, ioVBProject As Object)

    Dim iTotalExistingVBComponents As Integer
    Dim iTotalNonExistingVBComponents As Integer
    Dim oFiles As Collection
    Dim oFile As Collection
    Dim sAnswer As String
    Dim sFileNameWithoutExtension As String

    iTotalExistingVBComponents = 0
    iTotalNonExistingVBComponents = 0
    Set oFiles = GetVBFiles(AdjustFilePath(isFolderPath), ioVBProject.VBComponents, True)
    For Each oFile In oFiles
        Call GetVBFile(oFile, sFileNameWithoutExtension)
        If sFileNameWithoutExtension <> cCurrentModuleName Then
            If VBComponentExists(ioVBProject.VBComponents, sFileNameWithoutExtension) Then
                iTotalExistingVBComponents = iTotalExistingVBComponents + 1
            Else
                iTotalNonExistingVBComponents = iTotalNonExistingVBComponents + 1
            End If
        End If
    Next
    sAnswer = InputBox("IMPORT! That will add " & iTotalNonExistingVBComponents & " VB components and replace " & iTotalExistingVBComponents & " of them. Do you want to continue? (type ""YES"" for yes)", , "NO")
    If sAnswer <> "YES" Then
        err.Raise 64235, , "Procedure aborted by user"
    Else
        Call ImportVBProject(isFolderPath, isFileName, ioVBProject, True)
    End If

End Sub

Public Sub ImportVBProject(isFolderPath As String, isFileName As String, ioVBProject As Object, Optional ibReplaceAllVBComponents As Boolean)
    ' Import all files located in the folder of the current document, into modules, class modules and forms.

    Dim oVBProject As Object
    Dim oVBComponent As Object
    Dim sFolderPath As String
    Dim sFileExt As String
    Dim oFiles As Collection
    Dim oFile As Collection
    Dim sFileNameWithoutExtension As String

    If isFolderPath = "" Then
        err.Raise 64231, , "Document must be saved before import can start"
        Exit Sub
    End If

    Set oVBProject = ioVBProject
    If oVBProject Is Nothing Then
        Exit Sub
    End If

    sFolderPath = AdjustFilePath(isFolderPath)
    Set oFiles = GetVBFiles(sFolderPath, ioVBProject.VBComponents, True)

    For Each oFile In oFiles
        Call GetVBFile(oFile, sFileNameWithoutExtension, sFileExt)
        If sFileNameWithoutExtension <> cCurrentModuleName Then
            If VBComponentExists(ioVBProject.VBComponents, sFileNameWithoutExtension) Then
                ' Already exists
                If ibReplaceAllVBComponents = False Then
                    err.Raise 64234, , "Cannot import module '" & oVBComponent.Name & "' as it already exists!"
                End If
                ' Delete in separate procedure, otherwise it's not immediately deleted
                ' and the Import will create a new module with numeric suffix added to name
                ' (more details: https://stackoverflow.com/questions/19800184/vbcomponents-remove-doesnt-always-remove-module)
                Call RemoveVBComponent(oVBProject.VBComponents, ioVBProject.VBComponents(sFileNameWithoutExtension))
            End If
            Call oVBProject.VBComponents.Import(sFolderPath & sFileNameWithoutExtension & "." & sFileExt)
        End If
    Next

End Sub

Sub DeleteVBComponentFiles(isFolderPath As String, isFileName As String, ioVBProject As Object)
    ' Delete all files which correspond to the VB components of the current document/project

    Dim sFolderPath As String
    Dim oFiles As Collection
    Dim oFile As Collection
    Dim sFileNameWithoutExtension As String
    Dim oFSFile As Object

    If isFolderPath = "" Then
        err.Raise 64231, , "Document '" & isFileName & "' must be saved so that a folder can be checked"
    End If

    Set oFiles = GetVBFiles(AdjustFilePath(isFolderPath), ioVBProject.VBComponents, False)

    For Each oFile In oFiles
        Call GetVBFile(oFile, sFileNameWithoutExtension, , , , oFSFile)
        If VBComponentExists(ioVBProject.VBComponents, sFileNameWithoutExtension) Then
            Call oFSFile.Delete
        End If
    Next

End Sub

Public Sub RemoveVBComponent(ioVBComponents As Object, ioVBComponent As Object)

    Call ioVBComponents.Remove(ioVBComponent)

End Sub

Function GetFileExtension(iiVBComponentType As Integer) As String 'vbext_ComponentType

    Dim sExt As String

    Select Case iiVBComponentType
        Case 2, 100 'vbext_ct_ClassModule, vbext_ct_Document
            sExt = "cls"
        Case 3 'vbext_ct_MSForm
            sExt = "frm"
        Case 1 'vbext_ct_StdModule
            sExt = "bas"
        Case Else
            sExt = ""
    End Select

    GetFileExtension = sExt

End Function

Public Function VBComponentExists(ioVBComponents As Object, isName As String) As Boolean

    On Error GoTo err

    VBComponentExists = True
    IsObject ioVBComponents(isName)
    Exit Function
err:
    VBComponentExists = False

End Function

Public Function ExistsInCollection(ioColl As Collection, ivKey As Variant) As Boolean
    'https://stackoverflow.com/questions/137845/determining-whether-an-object-is-a-member-of-a-collection-in-vba
    On Error GoTo err
    ExistsInCollection = True
    IsObject (ioColl.Item(ivKey))
    Exit Function
err:
    ExistsInCollection = False
End Function

Public Sub GetVBFile( _
    ioCollFile As Collection, _
    Optional esFileNameWithoutExtension As String, _
    Optional esExtensionName As String, _
    Optional eiVBComponentType As Integer, _
    Optional ebExistsInCurrentVBProject As String, _
    Optional eoFSFile As Object) ' vbext_ComponentType

    esFileNameWithoutExtension = ioCollFile(1)
    esExtensionName = ioCollFile(2)
    eiVBComponentType = ioCollFile(3)
    ebExistsInCurrentVBProject = ioCollFile(4)
    Set eoFSFile = ioCollFile(5)

End Sub

Function GetVBFiles(isFolder As String, ioVBComponents As Object, Optional ibIncludeVBFilesNotPartOfCurrentVBProject As Boolean = False) As Collection

    Dim oCollFiles As Collection
    Dim oCollFile As Collection
    Dim oFileSystem As Object
    Dim oFSFolder As Object
    Dim oFSFileSet As Object
    Dim oFSFile As Object
    Dim iExt As Integer 'vbext_ComponentType
    Dim sFileNameWithoutExtension As String
    Dim bExistsInCurrentVBProject As Boolean

    Set oCollFiles = New Collection
    Set oFileSystem = CreateObject("Scripting.FileSystemObject")
    Set oFSFolder = oFileSystem.GetFolder(isFolder)
    Set oFSFileSet = oFSFolder.Files
    For Each oFSFile In oFSFileSet
        iExt = GetVBComponentType(oFSFile.Name)
        If iExt <> -1 Then
            sFileNameWithoutExtension = Left(oFSFile.Name, InStrRev(oFSFile.Name, ".") - 1)
            bExistsInCurrentVBProject = VBComponentExists(ioVBComponents, sFileNameWithoutExtension)
            If ibIncludeVBFilesNotPartOfCurrentVBProject _
                    Or bExistsInCurrentVBProject Then
                Set oCollFile = New Collection
                Call oCollFile.Add(sFileNameWithoutExtension) ' (1)
                Call oCollFile.Add(Mid(oFSFile.Name, InStrRev(oFSFile.Name, ".") + 1)) ' (2)
                Call oCollFile.Add(iExt) ' (3)
                Call oCollFile.Add(bExistsInCurrentVBProject) '(4)
                Call oCollFile.Add(oFSFile) '(5)
                Call oCollFiles.Add(oCollFile)
            End If
        End If
    Next

    Set GetVBFiles = oCollFiles

End Function

Function GetVBComponentType(isFileName As String) As Integer 'vbext_ComponentType

    Dim iExt As Integer 'vbext_ComponentType

    Select Case Right(isFileName, 4)
        Case ".bas":
            iExt = 1 'vbext_ct_StdModule
        Case ".cls":
            iExt = 2 'vbext_ct_ClassModule
        Case ".frm":
            iExt = 3 'vbext_ct_MSForm
        Case Else:
            iExt = -1
    End Select
    GetVBComponentType = iExt
End Function

Public Sub OpenAllWindows()

    Dim oVBComponent As Object

    For Each oVBComponent In Application.vbe.ActiveVBProject.VBComponents
        Call oVBComponent.Activate
    Next

End Sub

Public Sub CloseAllWindows()

    Dim oWindow As Window

    For Each oWindow In Application.vbe.ActiveVBProject.vbe.Windows
        On Error Resume Next
        If Not oWindow Is Application.vbe.ActiveWindow Then
            Select Case oWindow.Type
                Case vbext_wt_CodeWindow, vbext_wt_Designer
                    Call oWindow.Close
            End Select
        End If
    Next

End Sub

Public Sub DeleteMultipleVBComponentsWithOnlyOneConfirmationPopup(ioVBProject As Object)

    Dim sAnswer As Boolean
    Dim oVBComponent As Object

    sAnswer = InputBox("ALL VB components will be deleted EXCEPT Module1 and Module2. Do you want to continue? (type ""YES"" to continue)", , "NO")
    If sAnswer <> "YES" Then
        err.Raise 64235, , "Procedure aborted by user"
    End If

    For Each oVBComponent In ioVBProject.VBComponents
        Select Case oVBComponent.Name
            Case "Module1", "Module2":
            Case Else:
                Call ioVBProject.VBComponents.Remove(oVBComponent)
        End Select
    Next

End Sub

Function GetFolderPath(isFilePath As String) As String
    'Takes folder path out of a file path (e.g. C:\folder\file.bas -> C:\folder\)

    Dim pos As Integer
    pos = InStrRev(isFilePath, "\")
    If pos = 0 Then
        ' http(s)://.../... (OneDrive for instance)
        pos = InStrRev(isFilePath, "/")
    End If
    GetFolderPath = Left(isFilePath, pos)

End Function

Function AdjustFilePath(isFilePath As String) As String

    ' OneDrive
    AdjustFilePath = strOneDriveLocalFilePath(isFilePath)

End Function

Function GetNormalizedFolderPath(isFolderPath As String) As String

    Dim sFolderPath As String

    ' OneDrive
    sFolderPath = AdjustFilePath(isFolderPath)

    ' Add last character so that it's easier to build file path
    If InStr(sFolderPath, "\") > 0 Then
        If Right(sFolderPath, 1) <> "\" Then
            sFolderPath = sFolderPath & "\"
        End If
    Else
        If Right(sFolderPath, 1) <> "/" Then
            sFolderPath = sFolderPath & "/"
        End If
    End If

    GetNormalizedFolderPath = sFolderPath

End Function

Private Function strOneDriveLocalFilePath(sFolderPath As String) As String
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
            If InStr(1, sFolderPath, "my.sharepoint.com") <> 0 Then
                'find "/Documents" in string and replace everything before the end with OneDrive local path
                iDocumentsPosition = InStr(1, sFolderPath, "/Documents") + Len("/Documents") 'find "/Documents" position in file URL
                strFileEndPath = Mid(sFolderPath, iDocumentsPosition, Len(sFolderPath) - iDocumentsPosition + 1)  'get the ending file path without pointer in OneDrive
            Else
                'do nothing
            End If
            'for personal onedrive it looks like "https://d.docs.live.net/d7bbaa#######1/" & file.FullName, _
            '   by replacing "https.." with OneDrive local path obtained from registry we can get local file path
            If InStr(1, sFolderPath, "d.docs.live.net") <> 0 Then
                iSlashCount = 1
                i4thSlashPosition = 1
                Do Until iSlashCount > 4
                    i4thSlashPosition = InStr(i4thSlashPosition + 1, sFolderPath, "/")   'loop 4 times, looking for "/" after last found
                    iSlashCount = iSlashCount + 1
                Loop
                strFileEndPath = Mid(sFolderPath, i4thSlashPosition, Len(sFolderPath) - i4thSlashPosition + 1)  'get the ending file path without pointer in OneDrive
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
    If Not blnFileExist Then err.Raise 64230, , "File could not be found in any OneDrive folders"

    'clean up
    Set ShellScript = Nothing
    Set objFSO = Nothing
End Function



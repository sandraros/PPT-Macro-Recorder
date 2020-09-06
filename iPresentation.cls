VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "iPresentation"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Public Application As iApplication
Public AutoSaveOn As Boolean
'Public Broadcast As iBroadcast
Public BuiltInDocumentProperties As Object
'Public ChartDataPointTrack As Boolean
'Public Coauthoring As iCoauthoring
'Public ColorSchemes As iColorSchemes
'Public CommandBars As iCommandBars
'Public Container As iContainer
'Public ContentTypeProperties As iMetaProperties
Public CreateVideoStatus As PpMediaTaskStatus
Public CustomDocumentProperties As Object
'Public CustomerData As iCustomerData
'Public CustomXMLParts As iCustomXMLParts
Public DefaultLanguageID As MsoLanguageID
Public DefaultShape As iShape
'Public Designs As iDesigns
Public DisplayComments As MsoTriState
'Public DocumentInspectors As iDocumentInspectors
'Public DocumentLibraryVersions As iDocumentLibraryVersions
Public EncryptionProvider As String
Public EnvelopeVisible As MsoTriState
'Public ExtraColors As iExtraColors
Public FarEastLineBreakLanguage As MsoFarEastLineBreakLanguageID
Public FarEastLineBreakLevel As PpFarEastLineBreakLevel
Public Final As Boolean
'Public Fonts As iFonts
Public FullName As String
Public GridDistance As Single
'Public Guides As iGuides
'Public HandoutMaster As iMaster
Public HasHandoutMaster As Boolean
Public HasNotesMaster As Boolean
Public HasTitleMaster As MsoTriState
Public HasVBProject As Boolean
Public InMergeMode As Boolean
Public LayoutDirection As PpDirection
Public Name As String
Public NoLineBreakAfter As String
Public NoLineBreakBefore As String
'Public NotesMaster As iMaster
'Public PageSetup As iPageSetup
Public Parent As Object
Public Password As String
Public PasswordEncryptionAlgorithm As String
Public PasswordEncryptionFileProperties As Boolean
Public PasswordEncryptionKeyLength As Long
Public PasswordEncryptionProvider As String
Public Path As String
'Public Permission As iPermission
'Public PrintOptions As iPrintOptions
Public ReadOnly As MsoTriState
Public ReadOnlyRecommended As Boolean
Public RemovePersonalInformation As MsoTriState
'Public Research As iResearch
Public Saved As MsoTriState
'Public SectionProperties As iSectionProperties
'Public ServerPolicy As iServerPolicy
'Public SharedWorkspace As iSharedWorkspace
'Public Signatures As iSignatureSet
'Public SlideMaster As iMaster
Public Slides As iSlides
'Public SlideShowSettings As iSlideShowSettings
'Public SlideShowWindow As iSlideShowWindow
Public SnapToGrid As MsoTriState
'Public Sync As iSync
'Public Tags As iTags
Public TemplateName As String
'Public TitleMaster As iMaster
Public VBASigned As MsoTriState
'Public VBProject As iVBProject
'Public Windows As iDocumentWindows
Public WorkIdentity As String
Public WritePassword As String


'ActivePresentation.Parent


Public Function compare(propertyName As String, iPresentation As iPresentation, ByVal indent As Integer) As String

    On Error GoTo err_

    If IsCompared(Me) Then Exit Function

    code = ""

    With Me
        'code = code & .Application.compare("Application", iPresentation.Application, indent + 4)
        If .AutoSaveOn <> iPresentation.AutoSaveOn Then
            code = code & Space(indent) & ".AutoSaveOn = " & iPresentation.AutoSaveOn & Chr(13)
        End If
        If Not .BuiltInDocumentProperties Is iPresentation.BuiltInDocumentProperties Then
            code = code & Space(indent) & "'TODO BuiltInDocumentProperties" & Chr(13)
        End If
        If .CreateVideoStatus <> iPresentation.CreateVideoStatus Then
            code = code & Space(indent) & ".CreateVideoStatus = " & PpMediaTaskStatus(.CreateVideoStatus) & Chr(13)
        End If
        If Not .CustomDocumentProperties Is iPresentation.CustomDocumentProperties Then
            code = code & Space(indent) & "'TODO CustomDocumentProperties" & Chr(13)
        End If
        If .DefaultLanguageID <> iPresentation.DefaultLanguageID Then
            code = code & Space(indent) & ".DefaultLanguageID = " & MsoLanguageID(.DefaultLanguageID) & Chr(13)
        End If
        'TODO code = code & .DefaultShape.compare("DefaultShape", iPresentation.DefaultShape, indent + 4)
        If .DisplayComments <> iPresentation.DisplayComments Then
            code = code & Space(indent) & ".DisplayComments = " & MsoTriState(.DisplayComments) & Chr(13)
        End If
        If .EncryptionProvider <> iPresentation.EncryptionProvider Then
            code = code & Space(indent) & ".EncryptionProvider = " & iPresentation.EncryptionProvider & Chr(13)
        End If
        If .EnvelopeVisible <> iPresentation.EnvelopeVisible Then
            code = code & Space(indent) & ".EnvelopeVisible = " & MsoTriState(.EnvelopeVisible) & Chr(13)
        End If
        If .FarEastLineBreakLanguage <> iPresentation.FarEastLineBreakLanguage Then
            code = code & Space(indent) & ".FarEastLineBreakLanguage = " & MsoFarEastLineBreakLanguageID(.FarEastLineBreakLanguage) & Chr(13)
        End If
        If .FarEastLineBreakLevel <> iPresentation.FarEastLineBreakLevel Then
            code = code & Space(indent) & ".FarEastLineBreakLevel = " & PpFarEastLineBreakLevel(.FarEastLineBreakLevel) & Chr(13)
        End If
        If .Final <> iPresentation.Final Then
            code = code & Space(indent) & ".Final = " & iPresentation.Final & Chr(13)
        End If
        If .FullName <> iPresentation.FullName Then
            code = code & Space(indent) & ".FullName = " & iPresentation.FullName & Chr(13)
        End If
        If .GridDistance <> iPresentation.GridDistance Then
            code = code & Space(indent) & ".GridDistance = " & SingleToVBA(.GridDistance) & Chr(13)
        End If
        If .HasHandoutMaster <> iPresentation.HasHandoutMaster Then
            code = code & Space(indent) & ".HasHandoutMaster = " & iPresentation.HasHandoutMaster & Chr(13)
        End If
        If .HasNotesMaster <> iPresentation.HasNotesMaster Then
            code = code & Space(indent) & ".HasNotesMaster = " & iPresentation.HasNotesMaster & Chr(13)
        End If
        If .HasTitleMaster <> iPresentation.HasTitleMaster Then
            code = code & Space(indent) & ".HasTitleMaster = " & MsoTriState(.HasTitleMaster) & Chr(13)
        End If
        If .HasVBProject <> iPresentation.HasVBProject Then
            code = code & Space(indent) & ".HasVBProject = " & iPresentation.HasVBProject & Chr(13)
        End If
        If .InMergeMode <> iPresentation.InMergeMode Then
            code = code & Space(indent) & ".InMergeMode = " & iPresentation.InMergeMode & Chr(13)
        End If
        If .LayoutDirection <> iPresentation.LayoutDirection Then
            code = code & Space(indent) & ".LayoutDirection = " & PpDirection(.LayoutDirection) & Chr(13)
        End If
        If .Name <> iPresentation.Name Then
            code = code & Space(indent) & ".Name = " & iPresentation.Name & Chr(13)
        End If
        If .NoLineBreakAfter <> iPresentation.NoLineBreakAfter Then
            code = code & Space(indent) & ".NoLineBreakAfter = " & iPresentation.NoLineBreakAfter & Chr(13)
        End If
        If .NoLineBreakBefore <> iPresentation.NoLineBreakBefore Then
            code = code & Space(indent) & ".NoLineBreakBefore = " & iPresentation.NoLineBreakBefore & Chr(13)
        End If
        If Not .Parent Is iPresentation.Parent Then
            code = code & Space(indent) & "'TODO Parent"
        End If
        If .Password <> iPresentation.Password Then
            code = code & Space(indent) & ".Password = " & iPresentation.Password & Chr(13)
        End If
        If .PasswordEncryptionAlgorithm <> iPresentation.PasswordEncryptionAlgorithm Then
            code = code & Space(indent) & ".PasswordEncryptionAlgorithm = " & iPresentation.PasswordEncryptionAlgorithm & Chr(13)
        End If
        If .PasswordEncryptionFileProperties <> iPresentation.PasswordEncryptionFileProperties Then
            code = code & Space(indent) & ".PasswordEncryptionFileProperties = " & iPresentation.PasswordEncryptionFileProperties & Chr(13)
        End If
        If .PasswordEncryptionKeyLength <> iPresentation.PasswordEncryptionKeyLength Then
            code = code & Space(indent) & ".PasswordEncryptionKeyLength = " & LongToVBA(.PasswordEncryptionKeyLength) & Chr(13)
        End If
        If .PasswordEncryptionProvider <> iPresentation.PasswordEncryptionProvider Then
            code = code & Space(indent) & ".PasswordEncryptionProvider = " & iPresentation.PasswordEncryptionProvider & Chr(13)
        End If
        If .Path <> iPresentation.Path Then
            code = code & Space(indent) & ".Path = " & iPresentation.Path & Chr(13)
        End If
        If .ReadOnly <> iPresentation.ReadOnly Then
            code = code & Space(indent) & ".ReadOnly = " & MsoTriState(.ReadOnly) & Chr(13)
        End If
        If .ReadOnlyRecommended <> iPresentation.ReadOnlyRecommended Then
            code = code & Space(indent) & ".ReadOnlyRecommended = " & iPresentation.ReadOnlyRecommended & Chr(13)
        End If
        If .RemovePersonalInformation <> iPresentation.RemovePersonalInformation Then
            code = code & Space(indent) & ".RemovePersonalInformation = " & MsoTriState(.RemovePersonalInformation) & Chr(13)
        End If
        'If .Saved <> iPresentation.Saved Then
        '    code = code & Space(indent) & ".Saved = " & MsoTriState(.Saved) & Chr(13)
        'End If
        code = code & .Slides.compare(".Slides", iPresentation.Slides, indent + 4)
        If .SnapToGrid <> iPresentation.SnapToGrid Then
            code = code & Space(indent) & ".SnapToGrid = " & MsoTriState(.SnapToGrid) & Chr(13)
        End If
        If .TemplateName <> iPresentation.TemplateName Then
            code = code & Space(indent) & ".TemplateName = " & iPresentation.TemplateName & Chr(13)
        End If
        If .VBASigned <> iPresentation.VBASigned Then
            code = code & Space(indent) & ".VBASigned = " & MsoTriState(.VBASigned) & Chr(13)
        End If
        If .WorkIdentity <> iPresentation.WorkIdentity Then
            code = code & Space(indent) & ".WorkIdentity = " & iPresentation.WorkIdentity & Chr(13)
        End If
        If .WritePassword <> iPresentation.WritePassword Then
            code = code & Space(indent) & ".WritePassword = " & iPresentation.WritePassword & Chr(13)
        End If
    End With

    If code <> "" Then
        code = Space(indent - 4) & "With " & propertyName & Chr(13) _
            & code _
            & Space(indent - 4) & "End With" & Chr(13)
    End If

    compare = code

    Exit Function

err_:
    #If DEBUG_MODE = 1 Then
        Stop
    #Else
        err.Raise err.number 'rethrows with same source and description
    #End If

End Function

Public Sub init(iPresentation As Presentation)
    
    On Error GoTo err_
    
    With Me
        Set .Application = New_iApplication(iPresentation.Application)
        .AutoSaveOn = iPresentation.AutoSaveOn
        Set .BuiltInDocumentProperties = iPresentation.BuiltInDocumentProperties
        .CreateVideoStatus = iPresentation.CreateVideoStatus
        Set .CustomDocumentProperties = iPresentation.CustomDocumentProperties
        .DefaultLanguageID = iPresentation.DefaultLanguageID
        'Set .DefaultShape = New_iShape(iPresentation.DefaultShape)
        .DisplayComments = iPresentation.DisplayComments
        .EncryptionProvider = iPresentation.EncryptionProvider
        .EnvelopeVisible = iPresentation.EnvelopeVisible
        If iPresentation.FarEastLineBreakLevel = ppFarEastLineBreakLevelCustom Then
            .FarEastLineBreakLanguage = iPresentation.FarEastLineBreakLanguage
        End If
        .FarEastLineBreakLevel = iPresentation.FarEastLineBreakLevel
        .Final = iPresentation.Final
        .FullName = iPresentation.FullName
        .GridDistance = iPresentation.GridDistance
        .HasHandoutMaster = iPresentation.HasHandoutMaster
        .HasNotesMaster = iPresentation.HasNotesMaster
        .HasTitleMaster = iPresentation.HasTitleMaster
        .HasVBProject = iPresentation.HasVBProject
        .InMergeMode = iPresentation.InMergeMode
        .LayoutDirection = iPresentation.LayoutDirection
        .Name = iPresentation.Name
        .NoLineBreakAfter = iPresentation.NoLineBreakAfter
        .NoLineBreakBefore = iPresentation.NoLineBreakBefore
        Set .Parent = iPresentation.Parent
        .Password = iPresentation.Password
        .PasswordEncryptionAlgorithm = iPresentation.PasswordEncryptionAlgorithm
        .PasswordEncryptionFileProperties = iPresentation.PasswordEncryptionFileProperties
        .PasswordEncryptionKeyLength = iPresentation.PasswordEncryptionKeyLength
        .PasswordEncryptionProvider = iPresentation.PasswordEncryptionProvider
        .Path = iPresentation.Path
        .ReadOnly = iPresentation.ReadOnly
        .ReadOnlyRecommended = iPresentation.ReadOnlyRecommended
        .RemovePersonalInformation = iPresentation.RemovePersonalInformation
        .Saved = iPresentation.Saved
        Set .Slides = New_iSlides(iPresentation.Slides)
        .SnapToGrid = iPresentation.SnapToGrid
        .TemplateName = iPresentation.TemplateName
        .VBASigned = iPresentation.VBASigned
        .WorkIdentity = iPresentation.WorkIdentity
        .WritePassword = iPresentation.WritePassword
    End With

    Exit Sub

err_:
    #If DEBUG_MODE = 1 Then
        Stop
    #Else
        err.Raise err.number 'rethrows with same source and description
    #End If

End Sub


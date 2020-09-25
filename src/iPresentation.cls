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
Public defaultShape As iShape
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


Public Function Compare(ObjectName As String, iPresentation As iPresentation) As UDiff

    Dim oCode As UDiff

    On Error GoTo err_

    Set oCode = New_UDiff(ObjectName, Me, iPresentation)
    Set Compare = oCode

    If IsCompared(Me) Then Exit Function

    With Me
        'Call oCode.AddCode(.Application.compare("Application", iPresentation.Application)
        If .AutoSaveOn <> iPresentation.AutoSaveOn Then
            Call oCode.AddScalarProperty("AutoSaveOn", iPresentation.AutoSaveOn)
        End If
        'If Not .BuiltInDocumentProperties Is iPresentation.BuiltInDocumentProperties Then
        '    Call oCode.Add("'TODO BuiltInDocumentProperties")
        'End If
        If .CreateVideoStatus <> iPresentation.CreateVideoStatus Then
            Call oCode.AddScalarProperty("CreateVideoStatus", PpMediaTaskStatus(.CreateVideoStatus))
        End If
        'If Not .CustomDocumentProperties Is iPresentation.CustomDocumentProperties Then
        '    Call oCode.Add("'TODO CustomDocumentProperties")
        'End If
        If .DefaultLanguageID <> iPresentation.DefaultLanguageID Then
            Call oCode.AddScalarProperty("DefaultLanguageID", MsoLanguageID(.DefaultLanguageID))
        End If
        'TODO Call oCode.AddCode(.DefaultShape.compare("DefaultShape", iPresentation.DefaultShape)
        If .DisplayComments <> iPresentation.DisplayComments Then
            Call oCode.AddScalarProperty("DisplayComments", MsoTriState(.DisplayComments))
        End If
        If .EncryptionProvider <> iPresentation.EncryptionProvider Then
            Call oCode.AddScalarProperty("EncryptionProvider", iPresentation.EncryptionProvider)
        End If
        If .EnvelopeVisible <> iPresentation.EnvelopeVisible Then
            Call oCode.AddScalarProperty("EnvelopeVisible", MsoTriState(.EnvelopeVisible))
        End If
        If .FarEastLineBreakLanguage <> iPresentation.FarEastLineBreakLanguage Then
            Call oCode.AddScalarProperty("FarEastLineBreakLanguage", MsoFarEastLineBreakLanguageID(.FarEastLineBreakLanguage))
        End If
        If .FarEastLineBreakLevel <> iPresentation.FarEastLineBreakLevel Then
            Call oCode.AddScalarProperty("FarEastLineBreakLevel", PpFarEastLineBreakLevel(.FarEastLineBreakLevel))
        End If
        If .Final <> iPresentation.Final Then
            Call oCode.AddScalarProperty("Final", iPresentation.Final)
        End If
        If .FullName <> iPresentation.FullName Then
            Call oCode.AddScalarProperty("FullName", iPresentation.FullName)
        End If
        If .GridDistance <> iPresentation.GridDistance Then
            Call oCode.AddScalarProperty("GridDistance", SingleToVBA(.GridDistance))
        End If
        If .HasHandoutMaster <> iPresentation.HasHandoutMaster Then
            Call oCode.AddScalarProperty("HasHandoutMaster", iPresentation.HasHandoutMaster)
        End If
        If .HasNotesMaster <> iPresentation.HasNotesMaster Then
            Call oCode.AddScalarProperty("HasNotesMaster", iPresentation.HasNotesMaster)
        End If
        If .HasTitleMaster <> iPresentation.HasTitleMaster Then
            Call oCode.AddScalarProperty("HasTitleMaster", MsoTriState(.HasTitleMaster))
        End If
        If .HasVBProject <> iPresentation.HasVBProject Then
            Call oCode.AddScalarProperty("HasVBProject", iPresentation.HasVBProject)
        End If
        If .InMergeMode <> iPresentation.InMergeMode Then
            Call oCode.AddScalarProperty("InMergeMode", iPresentation.InMergeMode)
        End If
        If .LayoutDirection <> iPresentation.LayoutDirection Then
            Call oCode.AddScalarProperty("LayoutDirection", PpDirection(.LayoutDirection))
        End If
        If .Name <> iPresentation.Name Then
            Call oCode.AddScalarProperty("Name", iPresentation.Name)
        End If
        If .NoLineBreakAfter <> iPresentation.NoLineBreakAfter Then
            Call oCode.AddScalarProperty("NoLineBreakAfter", iPresentation.NoLineBreakAfter)
        End If
        If .NoLineBreakBefore <> iPresentation.NoLineBreakBefore Then
            Call oCode.AddScalarProperty("NoLineBreakBefore", iPresentation.NoLineBreakBefore)
        End If
        'If Not .Parent Is iPresentation.Parent Then
        '    Call oCode.Add("'TODO Parent")
        'End If
        If .Password <> iPresentation.Password Then
            Call oCode.AddScalarProperty("Password", iPresentation.Password)
        End If
        If .PasswordEncryptionAlgorithm <> iPresentation.PasswordEncryptionAlgorithm Then
            Call oCode.AddScalarProperty("PasswordEncryptionAlgorithm", iPresentation.PasswordEncryptionAlgorithm)
        End If
        If .PasswordEncryptionFileProperties <> iPresentation.PasswordEncryptionFileProperties Then
            Call oCode.AddScalarProperty("PasswordEncryptionFileProperties", iPresentation.PasswordEncryptionFileProperties)
        End If
        If .PasswordEncryptionKeyLength <> iPresentation.PasswordEncryptionKeyLength Then
            Call oCode.AddScalarProperty("PasswordEncryptionKeyLength", LongToVBA(.PasswordEncryptionKeyLength))
        End If
        If .PasswordEncryptionProvider <> iPresentation.PasswordEncryptionProvider Then
            Call oCode.AddScalarProperty("PasswordEncryptionProvider", iPresentation.PasswordEncryptionProvider)
        End If
        If .Path <> iPresentation.Path Then
            Call oCode.AddScalarProperty("Path", iPresentation.Path)
        End If
        If .ReadOnly <> iPresentation.ReadOnly Then
            Call oCode.AddScalarProperty("ReadOnly", MsoTriState(.ReadOnly))
        End If
        If .ReadOnlyRecommended <> iPresentation.ReadOnlyRecommended Then
            Call oCode.AddScalarProperty("ReadOnlyRecommended", iPresentation.ReadOnlyRecommended)
        End If
        If .RemovePersonalInformation <> iPresentation.RemovePersonalInformation Then
            Call oCode.AddScalarProperty("RemovePersonalInformation", MsoTriState(.RemovePersonalInformation))
        End If
        'If .Saved <> iPresentation.Saved Then
        '    Call oCode.AddScalarProperty("Saved", MsoTriState(.Saved))
        'End If
        Call oCode.AddCode(.Slides.Compare("Slides", iPresentation.Slides))
        If .SnapToGrid <> iPresentation.SnapToGrid Then
            Call oCode.AddScalarProperty("SnapToGrid", MsoTriState(.SnapToGrid))
        End If
        If .TemplateName <> iPresentation.TemplateName Then
            Call oCode.AddScalarProperty("TemplateName", iPresentation.TemplateName)
        End If
        If .VBASigned <> iPresentation.VBASigned Then
            Call oCode.AddScalarProperty("VBASigned", MsoTriState(.VBASigned))
        End If
        If .WorkIdentity <> iPresentation.WorkIdentity Then
            Call oCode.AddScalarProperty("WorkIdentity", iPresentation.WorkIdentity)
        End If
        If .WritePassword <> iPresentation.WritePassword Then
            Call oCode.AddScalarProperty("WritePassword", iPresentation.WritePassword)
        End If
    End With

    Call oCode.Wrap

    Set Compare = oCode

    Exit Function

err_:
    #If DEBUG_MODE = 1 Then
        Stop
    #Else
        err.Raise err.number 'rethrows with same source and description
    #End If

End Function

Public Sub init(iPresentation As presentation)

    On Error GoTo err_

    With Me
        Set .Application = New_iApplication(iPresentation.Application)
        .AutoSaveOn = iPresentation.AutoSaveOn
        Set .BuiltInDocumentProperties = iPresentation.BuiltInDocumentProperties
        .CreateVideoStatus = iPresentation.CreateVideoStatus
        Set .CustomDocumentProperties = iPresentation.CustomDocumentProperties
        .DefaultLanguageID = iPresentation.DefaultLanguageID
        Set .defaultShape = New_iShape(iPresentation.defaultShape)
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
        On Error Resume Next ' ReadOnlyRecommended don't exist in PPT 2002 but exist in PPT 2008
        .ReadOnlyRecommended = iPresentation.ReadOnlyRecommended
        On Error GoTo err_
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

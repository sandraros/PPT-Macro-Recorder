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
Public ChartDataPointTrack
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


'activepresentation.Application

Public Sub class_initialize()

    With Me
    End With

End Sub

Public Function compare(ByVal iColorFormat2 As iColorFormat2, indent As Integer) As String

    On Error GoTo err_

    code = ""
    With Me
    End With

    compare = code

    Exit Function

err_:
    #If DEBUG_MODE = 1 Then
        Stop
    #Else
        err.Raise err.number 'rethrows with same source and description
    #End If

End Function

Public Sub Init(iColorFormat As Office.ColorFormat)

    On Error GoTo err_

    If iColorFormat.Type = msoColorTypeMixed Then
        Exit Sub
    End If

    With Me
    End With

    Exit Sub

err_:
    #If DEBUG_MODE = 1 Then
        Stop
    #Else
        err.Raise err.number 'rethrows with same source and description
    #End If

End Sub


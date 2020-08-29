VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "iApplication"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Public Active As MsoTriState
Public ActiveEncryptionSession As Long
Public ActivePresentation As iPresentation
Public ActivePrinter As String
'Public ActiveProtectedViewWindow as iProtectedViewWindow
Public ActiveWindow As iDocumentWindow
'Public AddIns As iAddIns
'Public Assistance As iIAssistance
'Public AutoCorrect As iAutoCorrect
Public AutomationSecurity As MsoAutomationSecurity
Public Build As String
Public Caption As String
Public ChartDataPointTrack As Boolean
'Public COMAddIns As iCOMAddIns
'Public CommandBars as iCommandBars
Public Creator As Long
Public DisplayAlerts As PpAlertLevel
Public DisplayDocumentInformationPanel As Boolean
Public DisplayGridLines As MsoTriState
Public DisplayGuides As MsoTriState
Public FeatureInstall As MsoFeatureInstall
'Public FileConverters As iFileConverters
Public FileValidation As MsoFileValidationMode
Public Height As Single
Public IsSandboxed As Boolean
'Public LanguageSettings As iLanguageSettings
Public Left As Single
Public Name As String
'Public NewPresentation As iNewFile
Public OperatingSystem As String
'public Options as iOptions
Public Path As String
Public Presentations As iPresentations
Public ProductCode As String
'Public ProtectedViewWindows As iProtectedViewWindows
Public ShowStartupDialog As MsoTriState
Public ShowWindowsInTaskbar As MsoTriState
'public SlideShowWindows as iSlideShowWindows
'Public SmartArtColors As iSmartArtColors
'Public SmartArtLayouts As iSmartArtLayouts
'Public SmartArtQuickStyles As iSmartArtQuickStyles
Public Top As Single
'Public VBE As iVBE
Public Version As String
Public Visible As MsoTriState
Public Width As Single
Public Windows As iDocumentWindows
Public WindowState As PpWindowState



'Application.WindowState

Public Sub class_initialize()

    With Me
        Set .ActivePresentation = New iPresentation
        Set .ActiveWindow = New iDocumentWindow
        Set .Presentations = New iPresentations
        Set .Windows = New iDocumentWindows
    End With

End Sub

Public Function compare(ByVal iApplication As iApplication, indent As Integer) As String

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

Public Sub Init(iApplication As Application)

    On Error GoTo err_

    With Me
        .Active = iApplication.Active
        .ActiveEncryptionSession = iApplication.ActiveEncryptionSession
        Call .ActivePresentation.Init(iApplication.ActivePresentation)
        .ActivePrinter = iApplication.ActivePrinter
        Call .ActiveWindow.Init(iApplication.ActiveWindow)
        .AutomationSecurity = iApplication.AutomationSecurity
        .Build = iApplication.Build
        .Caption = iApplication.Caption
        .ChartDataPointTrack = iApplication.ChartDataPointTrack
        .Creator = iApplication.Creator
        .DisplayAlerts = iApplication.DisplayAlerts
        .DisplayDocumentInformationPanel = iApplication.DisplayDocumentInformationPanel
        .DisplayGridLines = iApplication.DisplayGridLines
        .DisplayGuides = iApplication.DisplayGuides
        .FeatureInstall = iApplication.FeatureInstall
        .FileValidation = iApplication.FileValidation
        .Height = iApplication.Height
        .IsSandboxed = iApplication.IsSandboxed
        .Left = iApplication.Left
        .Name = iApplication.Name
        .OperatingSystem = iApplication.OperatingSystem
        .Path = iApplication.Path
        Call .Presentations.Init(iApplication.Presentations)
        .ProductCode = iApplication.ProductCode
        .ShowStartupDialog = iApplication.ShowStartupDialog
        .ShowWindowsInTaskbar = iApplication.ShowWindowsInTaskbar
        .Top = iApplication.Top
        .Version = iApplication.Version
        .Visible = iApplication.Visible
        .Width = iApplication.Width
        Call .Windows.Init(iApplication.Windows)
        .WindowState = iApplication.WindowState
    End With

    Exit Sub

err_:
    #If DEBUG_MODE = 1 Then
        Stop
    #Else
        err.Raise err.number 'rethrows with same source and description
    #End If

End Sub


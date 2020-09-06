VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "iApplication"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Public Value As String
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


'Application.Path


Sub Class_Initialize()
    Value = "hhh"
End Sub

Public Function compare(propertyName As String, iApplication As iApplication, indent As Integer) As String

    On Error GoTo err_

    If IsCompared(Me) Then Exit Function

    code = ""

    With Me
        If .Active <> iApplication.Active Then
            code = code & Space(indent) & ".Active = " & MsoTriState(.Active) & Chr(13)
        End If
        If .ActiveEncryptionSession <> iApplication.ActiveEncryptionSession Then
            code = code & Space(indent) & ".ActiveEncryptionSession = " & LongToVBA(.ActiveEncryptionSession) & Chr(13)
        End If
        code = code & .ActivePresentation.compare(".ActivePresentation", iApplication.ActivePresentation, indent + 4)
        If .ActivePrinter <> iApplication.ActivePrinter Then
            code = code & Space(indent) & ".ActivePrinter = " & iApplication.ActivePrinter & Chr(13)
        End If
        code = code & .ActiveWindow.compare(".ActiveWindow", iApplication.ActiveWindow, indent + 4)
        If .AutomationSecurity <> iApplication.AutomationSecurity Then
            code = code & Space(indent) & ".AutomationSecurity = " & MsoAutomationSecurity(.AutomationSecurity) & Chr(13)
        End If
        If .Build <> iApplication.Build Then
            code = code & Space(indent) & ".Build = " & iApplication.Build & Chr(13)
        End If
        If .Caption <> iApplication.Caption Then
            code = code & Space(indent) & ".Caption = " & iApplication.Caption & Chr(13)
        End If
        If .ChartDataPointTrack <> iApplication.ChartDataPointTrack Then
            code = code & Space(indent) & ".ChartDataPointTrack = " & iApplication.ChartDataPointTrack & Chr(13)
        End If
        If .Creator <> iApplication.Creator Then
            code = code & Space(indent) & ".Creator = " & LongToVBA(.Creator) & Chr(13)
        End If
        If .DisplayAlerts <> iApplication.DisplayAlerts Then
            code = code & Space(indent) & ".DisplayAlerts = " & PpAlertLevel(.DisplayAlerts) & Chr(13)
        End If
        If .DisplayDocumentInformationPanel <> iApplication.DisplayDocumentInformationPanel Then
            code = code & Space(indent) & ".DisplayDocumentInformationPanel = " & iApplication.DisplayDocumentInformationPanel & Chr(13)
        End If
        If .DisplayGridLines <> iApplication.DisplayGridLines Then
            code = code & Space(indent) & ".DisplayGridLines = " & MsoTriState(.DisplayGridLines) & Chr(13)
        End If
        If .DisplayGuides <> iApplication.DisplayGuides Then
            code = code & Space(indent) & ".DisplayGuides = " & MsoTriState(.DisplayGuides) & Chr(13)
        End If
        If .FeatureInstall <> iApplication.FeatureInstall Then
            code = code & Space(indent) & ".FeatureInstall = " & MsoFeatureInstall(.FeatureInstall) & Chr(13)
        End If
        If .FileValidation <> iApplication.FileValidation Then
            code = code & Space(indent) & ".FileValidation = " & MsoFileValidationMode(.FileValidation) & Chr(13)
        End If
        If .Height <> iApplication.Height Then
            code = code & Space(indent) & ".Height = " & SingleToVBA(.Height) & Chr(13)
        End If
        If .IsSandboxed <> iApplication.IsSandboxed Then
            code = code & Space(indent) & ".IsSandboxed = " & iApplication.IsSandboxed & Chr(13)
        End If
        If .Left <> iApplication.Left Then
            code = code & Space(indent) & ".Left = " & SingleToVBA(.Left) & Chr(13)
        End If
        If .Name <> iApplication.Name Then
            code = code & Space(indent) & ".Name = " & iApplication.Name & Chr(13)
        End If
        If .OperatingSystem <> iApplication.OperatingSystem Then
            code = code & Space(indent) & ".OperatingSystem = " & iApplication.OperatingSystem & Chr(13)
        End If
        If .Path <> iApplication.Path Then
            code = code & Space(indent) & ".Path = " & iApplication.Path & Chr(13)
        End If
        code = code & .Presentations.compare(".Presentations", iApplication.Presentations, indent + 4)
        If .ProductCode <> iApplication.ProductCode Then
            code = code & Space(indent) & ".ProductCode = " & iApplication.ProductCode & Chr(13)
        End If
        If .ShowStartupDialog <> iApplication.ShowStartupDialog Then
            code = code & Space(indent) & ".ShowStartupDialog = " & MsoTriState(.ShowStartupDialog) & Chr(13)
        End If
        'If .ShowWindowsInTaskbar <> iApplication.ShowWindowsInTaskbar Then
        '    code = code & ".ShowWindowsInTaskbar = " & MsoTriState(.ShowWindowsInTaskbar) & Chr(13)
        'End If
        If .Top <> iApplication.Top Then
            code = code & Space(indent) & ".Top = " & SingleToVBA(.Top) & Chr(13)
        End If
        If .Version <> iApplication.Version Then
            code = code & Space(indent) & ".Version = " & iApplication.Version & Chr(13)
        End If
        If .Visible <> iApplication.Visible Then
            code = code & Space(indent) & ".Visible = " & MsoTriState(.Visible) & Chr(13)
        End If
        If .Width <> iApplication.Width Then
            code = code & Space(indent) & ".Width = " & SingleToVBA(.Width) & Chr(13)
        End If
        code = code & .Windows.compare(".Windows", iApplication.Windows, indent + 4)
        If .WindowState <> iApplication.WindowState Then
            code = code & Space(indent) & ".WindowState = " & PpWindowState(.WindowState) & Chr(13)
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

Public Sub init(iApplication As Application)

    On Error GoTo err_

    With Me
        .Active = iApplication.Active
        .ActiveEncryptionSession = iApplication.ActiveEncryptionSession
        Set .ActivePresentation = New_iPresentation(iApplication.ActivePresentation)
        .ActivePrinter = iApplication.ActivePrinter
        Set .ActiveWindow = New_iDocumentWindow(iApplication.ActiveWindow)
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
        Set .Presentations = New_iPresentations(iApplication.Presentations)
        .ProductCode = iApplication.ProductCode
        .ShowStartupDialog = iApplication.ShowStartupDialog
        '.ShowWindowsInTaskbar = iApplication.ShowWindowsInTaskbar
        .Top = iApplication.Top
        .Version = iApplication.Version
        .Visible = iApplication.Visible
        .Width = iApplication.Width
        Set .Windows = New_iDocumentWindows(iApplication.Windows)
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


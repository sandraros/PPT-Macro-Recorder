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

Public Function compare(propertyName As String, iApplication As iApplication) As String

    On Error GoTo err_

    If IsCompared(Me) Then Exit Function

    Dim code As String
    code = ""

    With Me
        If .Active <> iApplication.Active Then
            Call addCode(code, ".Active = " & MsoTriState(.Active))
        End If
        If .ActiveEncryptionSession <> iApplication.ActiveEncryptionSession Then
            Call addCode(code, ".ActiveEncryptionSession = " & LongToVBA(.ActiveEncryptionSession))
        End If
        Call addCode(code, .ActivePresentation.compare(".ActivePresentation", iApplication.ActivePresentation))
        If .ActivePrinter <> iApplication.ActivePrinter Then
            Call addCode(code, ".ActivePrinter = " & iApplication.ActivePrinter)
        End If
        Call addCode(code, .ActiveWindow.compare(".ActiveWindow", iApplication.ActiveWindow))
        If .AutomationSecurity <> iApplication.AutomationSecurity Then
            Call addCode(code, ".AutomationSecurity = " & MsoAutomationSecurity(.AutomationSecurity))
        End If
        If .Build <> iApplication.Build Then
            Call addCode(code, ".Build = " & iApplication.Build)
        End If
        If .Caption <> iApplication.Caption Then
            Call addCode(code, ".Caption = " & iApplication.Caption)
        End If
        If .ChartDataPointTrack <> iApplication.ChartDataPointTrack Then
            Call addCode(code, ".ChartDataPointTrack = " & iApplication.ChartDataPointTrack)
        End If
        If .Creator <> iApplication.Creator Then
            Call addCode(code, ".Creator = " & LongToVBA(.Creator))
        End If
        If .DisplayAlerts <> iApplication.DisplayAlerts Then
            Call addCode(code, ".DisplayAlerts = " & PpAlertLevel(.DisplayAlerts))
        End If
        If .DisplayDocumentInformationPanel <> iApplication.DisplayDocumentInformationPanel Then
            Call addCode(code, ".DisplayDocumentInformationPanel = " & iApplication.DisplayDocumentInformationPanel)
        End If
        If .DisplayGridLines <> iApplication.DisplayGridLines Then
            Call addCode(code, ".DisplayGridLines = " & MsoTriState(.DisplayGridLines))
        End If
        If .DisplayGuides <> iApplication.DisplayGuides Then
            Call addCode(code, ".DisplayGuides = " & MsoTriState(.DisplayGuides))
        End If
        If .FeatureInstall <> iApplication.FeatureInstall Then
            Call addCode(code, ".FeatureInstall = " & MsoFeatureInstall(.FeatureInstall))
        End If
        If .FileValidation <> iApplication.FileValidation Then
            Call addCode(code, ".FileValidation = " & MsoFileValidationMode(.FileValidation))
        End If
        If .Height <> iApplication.Height Then
            Call addCode(code, ".Height = " & SingleToVBA(.Height))
        End If
        If .IsSandboxed <> iApplication.IsSandboxed Then
            Call addCode(code, ".IsSandboxed = " & iApplication.IsSandboxed)
        End If
        If .Left <> iApplication.Left Then
            Call addCode(code, ".Left = " & SingleToVBA(.Left))
        End If
        If .Name <> iApplication.Name Then
            Call addCode(code, ".Name = " & iApplication.Name)
        End If
        If .OperatingSystem <> iApplication.OperatingSystem Then
            Call addCode(code, ".OperatingSystem = " & iApplication.OperatingSystem)
        End If
        If .Path <> iApplication.Path Then
            Call addCode(code, ".Path = " & iApplication.Path)
        End If
        Call addCode(code, .Presentations.compare(".Presentations", iApplication.Presentations))
        If .ProductCode <> iApplication.ProductCode Then
            Call addCode(code, ".ProductCode = " & iApplication.ProductCode)
        End If
        If .ShowStartupDialog <> iApplication.ShowStartupDialog Then
            Call addCode(code, ".ShowStartupDialog = " & MsoTriState(.ShowStartupDialog))
        End If
        'If .ShowWindowsInTaskbar <> iApplication.ShowWindowsInTaskbar Then
        '    Call addCode(code, ".ShowWindowsInTaskbar = " & MsoTriState(.ShowWindowsInTaskbar))
        'End If
        If .Top <> iApplication.Top Then
            Call addCode(code, ".Top = " & SingleToVBA(.Top))
        End If
        If .Version <> iApplication.Version Then
            Call addCode(code, ".Version = " & iApplication.Version)
        End If
        If .Visible <> iApplication.Visible Then
            Call addCode(code, ".Visible = " & MsoTriState(.Visible))
        End If
        If .Width <> iApplication.Width Then
            Call addCode(code, ".Width = " & SingleToVBA(.Width))
        End If
        Call addCode(code, .Windows.compare(".Windows", iApplication.Windows))
        If .WindowState <> iApplication.WindowState Then
            Call addCode(code, ".WindowState = " & PpWindowState(.WindowState))
        End If
    End With

    Call Utility.WrapCode(propertyName, code)

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


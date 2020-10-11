VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "iDocumentWindow"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Public Active As MsoTriState
'public ActivePane as iPane
Public Application As iApplication
Public BlackAndWhite As MsoTriState
Public Caption As String
Public Height As Single
Public Left As Single
'Public Panes As iPanes
Public Parent As Object
Public Presentation As iPresentation
Public Selection As iSelection
Public SplitHorizontal As Long
Public SplitVertical As Long
Public Top As Single
'Public View As iView
Public ViewType As PpViewType
Public Width As Single
Public WindowState As PpWindowState


'ActivePresentation.Windows(1).WindowState


Public Function MR_Compare(iDocumentWindow As iDocumentWindow) As MR_Diff

    Dim oDiff As MR_Diff

    On Error GoTo err_

    Set oDiff = New_MR_Diff(Me, iDocumentWindow)
    Set MR_Compare = oDiff

    If IsCompared(Me) Then Exit Function

    With Me
        If .BlackAndWhite <> iDocumentWindow.BlackAndWhite Then
            Call oDiff.AddScalarProperty("BlackAndWhite", iDocumentWindow.BlackAndWhite)
        End If
        If .Caption <> iDocumentWindow.Caption Then
            Call oDiff.AddScalarProperty("Caption", iDocumentWindow.Caption)
        End If
        If .Height <> iDocumentWindow.Height Then
            Call oDiff.AddScalarProperty("Height", iDocumentWindow.Height)
        End If
        If .Left <> iDocumentWindow.Left Then
            Call oDiff.AddScalarProperty("Left", iDocumentWindow.Left)
        End If
        Call oDiff.AddDiff("Presentation", .Presentation.MR_Compare(iDocumentWindow.Presentation))
        Call oDiff.AddDiff("Selection", .Selection.MR_Compare(iDocumentWindow.Selection))
        If .SplitHorizontal <> iDocumentWindow.SplitHorizontal Then
            Call oDiff.AddScalarProperty("SplitHorizontal", iDocumentWindow.SplitHorizontal)
        End If
        If .SplitVertical <> iDocumentWindow.SplitVertical Then
            Call oDiff.AddScalarProperty("SplitVertical", iDocumentWindow.SplitVertical)
        End If
        If .Top <> iDocumentWindow.Top Then
            Call oDiff.AddScalarProperty("Top", iDocumentWindow.Top)
        End If
        If .ViewType <> iDocumentWindow.ViewType Then
            Call oDiff.AddScalarProperty("ViewType", iDocumentWindow.ViewType)
        End If
        If .Width <> iDocumentWindow.Width Then
            Call oDiff.AddScalarProperty("Width", iDocumentWindow.Width)
        End If
        If .WindowState <> iDocumentWindow.WindowState Then
            Call oDiff.AddScalarProperty("WindowState", iDocumentWindow.WindowState)
        End If
    End With

    Call oDiff.Wrap

    Set MR_Compare = oDiff

    Exit Function

err_:
    #If DEBUG_MODE = 1 Then
        Stop
    #Else
        err.Raise err.number 'rethrows with same source and description
    #End If

End Function

Public Sub init(iDocumentWindow As DocumentWindow)

    On Error GoTo err_

    With Me
        .Active = iDocumentWindow.Active
        Set .Application = New_iApplication(iDocumentWindow.Application)
        .BlackAndWhite = iDocumentWindow.BlackAndWhite
        .Caption = iDocumentWindow.Caption
        .Height = iDocumentWindow.Height
        .Left = iDocumentWindow.Left
        Set .Parent = iDocumentWindow.Parent
        Set .Presentation = New_iPresentation(iDocumentWindow.Presentation)
        Set .Selection = New_iSelection(iDocumentWindow.Selection)
        .SplitHorizontal = iDocumentWindow.SplitHorizontal
        .SplitVertical = iDocumentWindow.SplitVertical
        .Top = iDocumentWindow.Top
        .ViewType = iDocumentWindow.ViewType
        .Width = iDocumentWindow.Width
        .WindowState = iDocumentWindow.WindowState
    End With

    Exit Sub

err_:
    #If DEBUG_MODE = 1 Then
        Stop
    #Else
        err.Raise err.number 'rethrows with same source and description
    #End If

End Sub

Sub DefaultValues()

    On Error GoTo err_

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

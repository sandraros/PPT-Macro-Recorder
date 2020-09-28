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
Public presentation As iPresentation
Public Selection As iSelection
Public SplitHorizontal As Long
Public SplitVertical As Long
Public Top As Single
'Public View As iView
Public ViewType As PpViewType
Public Width As Single
Public WindowState As PpWindowState


'ActivePresentation.Windows(1).WindowState


Public Function MR_Compare(ObjectName As String, iDocumentWindow As iDocumentWindow) As MR_Diff

    Dim oDiff As MR_Diff

    On Error GoTo err_

    Set oDiff = New_MR_Diff(ObjectName, Me, iDocumentWindow)
    Set MR_Compare = oDiff

    If IsCompared(Me) Then Exit Function

    With Me
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
        Set .presentation = New_iPresentation(iDocumentWindow.presentation)
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

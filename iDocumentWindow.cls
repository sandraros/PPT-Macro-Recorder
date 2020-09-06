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


Public Function compare(propertyName As String, iDocumentWindow As iDocumentWindow, indent As Integer) As String

    On Error GoTo err_

    If IsCompared(Me) Then Exit Function

    code = ""
 
    With Me
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



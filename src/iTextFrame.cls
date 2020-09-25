VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "iTextFrame"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'OBSOLETE - TEXTFRAME SUPERSEDED WITH TEXTFRAME2

Public AutoSize As PpAutoSize
Public Creator As Long
Public HasText As MsoTriState
Public HorizontalAnchor As MsoHorizontalAnchor
Public MarginBottom As Single
Public MarginLeft As Single
Public MarginRight As Single
Public MarginTop As Single
Public Orientation As MsoTextOrientation
'Public Ruler As iRuler
Public TextRange As iTextRange
Public VerticalAnchor As MsoVerticalAnchor
Public WordWrap As MsoTriState


'ActiveWindow.Selection.ShapeRange.TextFrame.WordWrap


Public Function Compare(ObjectName As String, iTextFrame As iTextFrame) As UDiff

    Dim oCode As UDiff

    On Error GoTo err_

    Set oCode = New_UDiff(ObjectName, Me, iTextFrame)
    Set Compare = oCode

    If IsCompared(Me) Then Exit Function

    With Me
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

Public Sub init(iTextFrame As TextFrame)

    On Error GoTo err_

    belongsToDefaultShape = (TypeName(iTextFrame.Parent.Parent) = "Presentation")

    With Me
        .AutoSize = iTextFrame.AutoSize
        .Creator = iTextFrame.Creator
        If belongsToDefaultShape = False Then
            .HasText = iTextFrame.HasText
        End If
        .HorizontalAnchor = iTextFrame.HorizontalAnchor
        .MarginBottom = iTextFrame.MarginBottom
        .MarginLeft = iTextFrame.MarginLeft
        .MarginRight = iTextFrame.MarginRight
        .MarginTop = iTextFrame.MarginTop
        .Orientation = iTextFrame.Orientation
        If belongsToDefaultShape = False Then
            Set .TextRange = New_iTextRange(iTextFrame.TextRange)
        Else
            Set .TextRange = New iTextRange
        End If
        .VerticalAnchor = iTextFrame.VerticalAnchor
        .WordWrap = iTextFrame.WordWrap
    End With

    Exit Sub

err_:
    #If DEBUG_MODE = 1 Then
        Stop
    #Else
        err.Raise err.number 'rethrows with same source and description
    #End If

End Sub

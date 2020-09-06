VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "iTextFrame2"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Public Application As Object
Public AutoSize As MsoAutoSize
'Public Column As iTextColumn2
Public Creator As Long
Public HasText As MsoTriState
Public HorizontalAnchor As MsoHorizontalAnchor
Public MarginBottom As Single
Public MarginLeft As Single
Public MarginRight As Single
Public MarginTop As Single
'Public NoTextRotation As MsoTriState
Public Orientation As MsoTextOrientation
Public Parent As Object
Public PathFormat As MsoPathFormat
'Public Ruler As iRuler2
Public TextRange As iTextRange2
'Public ThreeD As iThreeDFormat
Public VerticalAnchor As MsoVerticalAnchor
Public WarpFormat As MsoWarpFormat
Public WordArtFormat As MsoPresetTextEffect
Public WordWrap As MsoTriState


'ActiveWindow.Selection.ShapeRange.TextFrame2.WordWrap


Public Function compare(propertyName As String, iTextFrame2 As iTextFrame2, indent As Integer) As String

    On Error GoTo err_

    If IsCompared(Me) Then Exit Function

    code = ""

    With Me
        'If .Application <> iTextFrame2.Application Then
        '    code = code & ".Application = " & iTextFrame2.Application & Chr(13)
        'End If
        If .AutoSize <> iTextFrame2.AutoSize Then
            code = code & Space(indent) & ".AutoSize = " & MsoAutoSize(.AutoSize) & Chr(13)
        End If
        'If .Creator <> iTextFrame2.Creator Then
        '    code = code & Space(indent) & ".Creator = " & LongToVBA(.Creator) & Chr(13)
        'End If
        'If .HasText <> iTextFrame2.HasText Then
        '    code = code & Space(indent) & ".HasText = " & MsoTriState(.HasText) & Chr(13)
        'End If
        If .HorizontalAnchor <> iTextFrame2.HorizontalAnchor Then
            code = code & Space(indent) & ".HorizontalAnchor = " & MsoHorizontalAnchor(.HorizontalAnchor) & Chr(13)
        End If
        If .MarginBottom <> iTextFrame2.MarginBottom Then
            code = code & Space(indent) & ".MarginBottom = " & SingleToVBA(.MarginBottom) & Chr(13)
        End If
        If .MarginLeft <> iTextFrame2.MarginLeft Then
            code = code & Space(indent) & ".MarginLeft = " & SingleToVBA(.MarginLeft) & Chr(13)
        End If
        If .MarginRight <> iTextFrame2.MarginRight Then
            code = code & Space(indent) & ".MarginRight = " & SingleToVBA(.MarginRight) & Chr(13)
        End If
        If .MarginTop <> iTextFrame2.MarginTop Then
            code = code & Space(indent) & ".MarginTop = " & SingleToVBA(.MarginTop) & Chr(13)
        End If
        If .Orientation <> iTextFrame2.Orientation Then
            code = code & Space(indent) & ".Orientation = " & MsoTextOrientation(.Orientation) & Chr(13)
        End If
        If Not .Parent Is iTextFrame2.Parent Then
            code = code & Space(indent) & "'TODO Parent" & Chr(13)
        End If
        If .PathFormat <> iTextFrame2.PathFormat Then
            code = code & Space(indent) & ".PathFormat = " & MsoPathFormat(.PathFormat) & Chr(13)
        End If
        code = code & .TextRange.compare(".TextRange", iTextFrame2.TextRange, indent + 4)
        If .VerticalAnchor <> iTextFrame2.VerticalAnchor Then
            code = code & Space(indent) & ".VerticalAnchor = " & MsoVerticalAnchor(.VerticalAnchor) & Chr(13)
        End If
        If .WarpFormat <> iTextFrame2.WarpFormat Then
            code = code & Space(indent) & ".WarpFormat = " & MsoWarpFormat(.WarpFormat) & Chr(13)
        End If
        If .WordArtFormat <> iTextFrame2.WordArtFormat Then
            code = code & Space(indent) & "'.WordArtFormat = " & MsoPresetTextEffect(.WordArtFormat) & Chr(13)
        End If
        If .WordWrap <> iTextFrame2.WordWrap Then
            code = code & Space(indent) & ".WordWrap = " & MsoTriState(.WordWrap) & Chr(13)
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

Public Sub init(iTextFrame2 As TextFrame2)

    On Error GoTo err_

    With Me
        Set .Application = iTextFrame2.Application
        .AutoSize = iTextFrame2.AutoSize
        .Creator = iTextFrame2.Creator
        .HasText = iTextFrame2.HasText
        .HorizontalAnchor = iTextFrame2.HorizontalAnchor
        .MarginBottom = iTextFrame2.MarginBottom
        .MarginLeft = iTextFrame2.MarginLeft
        .MarginRight = iTextFrame2.MarginRight
        .MarginTop = iTextFrame2.MarginTop
        '??? .NoTextRotation = iTextFrame2.NoTextRotation ' Returns or sets a value that specifies if the text on a shape is rotated if the shape itself is being rotated. Read/write.
        .Orientation = iTextFrame2.Orientation
        Set .Parent = iTextFrame2.Parent
        .PathFormat = iTextFrame2.PathFormat
        '??? If iTextFrame2.HasText Then
        Set .TextRange = New_iTextRange2(iTextFrame2.TextRange)
        '??? Else
        'Set .TextRange = New iTextRange2
        '??? End If
        .VerticalAnchor = iTextFrame2.VerticalAnchor
        .WarpFormat = iTextFrame2.WarpFormat
        .WordArtFormat = iTextFrame2.WordArtFormat
        .WordWrap = iTextFrame2.WordWrap
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
    With Me
        Set .TextRange = New iTextRange2
        Call .TextRange.DefaultValues
    End With
End Sub


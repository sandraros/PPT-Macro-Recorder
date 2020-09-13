VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "iTextFrame2"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Public AutoSize As MsoAutoSize
'Public Column As iTextColumn2
Public HasText As MsoTriState
Public HorizontalAnchor As MsoHorizontalAnchor
Public MarginBottom As Single
Public MarginLeft As Single
Public MarginRight As Single
Public MarginTop As Single
'Public NoTextRotation As MsoTriState
Public Orientation As MsoTextOrientation
Public PathFormat As MsoPathFormat
'Public Ruler As iRuler2
Public TextRange As iTextRange2
'Public ThreeD As iThreeDFormat
Public VerticalAnchor As MsoVerticalAnchor
Public WarpFormat As MsoWarpFormat
Public WordArtFormat As MsoPresetTextEffect
Public WordWrap As MsoTriState


'ActiveWindow.Selection.ShapeRange.TextFrame2.WordWrap


Public Function compare(propertyName As String, iTextFrame2 As iTextFrame2) As String

    On Error GoTo err_

    If IsCompared(Me) Then Exit Function

    Dim code As String
    code = ""

    With Me
        If .AutoSize <> iTextFrame2.AutoSize Then
            Call addCode(code, ".AutoSize = " & MsoAutoSize(.AutoSize))
        End If
        'If .HasText <> iTextFrame2.HasText Then
        '    Call addCode(code, ".HasText = " & MsoTriState(.HasText))
        'End If
        If .HorizontalAnchor <> iTextFrame2.HorizontalAnchor Then
            Call addCode(code, ".HorizontalAnchor = " & MsoHorizontalAnchor(.HorizontalAnchor))
        End If
        If .MarginBottom <> iTextFrame2.MarginBottom Then
            Call addCode(code, ".MarginBottom = " & SingleToVBA(.MarginBottom))
        End If
        If .MarginLeft <> iTextFrame2.MarginLeft Then
            Call addCode(code, ".MarginLeft = " & SingleToVBA(.MarginLeft))
        End If
        If .MarginRight <> iTextFrame2.MarginRight Then
            Call addCode(code, ".MarginRight = " & SingleToVBA(.MarginRight))
        End If
        If .MarginTop <> iTextFrame2.MarginTop Then
            Call addCode(code, ".MarginTop = " & SingleToVBA(.MarginTop))
        End If
        If .Orientation <> iTextFrame2.Orientation Then
            Call addCode(code, ".Orientation = " & MsoTextOrientation(.Orientation))
        End If
        If .PathFormat <> iTextFrame2.PathFormat Then
            Call addCode(code, ".PathFormat = " & MsoPathFormat(.PathFormat))
        End If
        Call addCode(code, .TextRange.compare(".TextRange", iTextFrame2.TextRange))
        If .VerticalAnchor <> iTextFrame2.VerticalAnchor Then
            Call addCode(code, ".VerticalAnchor = " & MsoVerticalAnchor(.VerticalAnchor))
        End If
        If .WarpFormat <> iTextFrame2.WarpFormat Then
            Call addCode(code, ".WarpFormat = " & MsoWarpFormat(.WarpFormat))
        End If
        If .WordArtFormat <> iTextFrame2.WordArtFormat Then
            Call addCode(code, "'.WordArtFormat = " & MsoPresetTextEffect(.WordArtFormat))
        End If
        If .WordWrap <> iTextFrame2.WordWrap Then
            Call addCode(code, ".WordWrap = " & MsoTriState(.WordWrap))
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

Public Sub init(iTextFrame2 As TextFrame2)

    On Error GoTo err_

    With Me
        .AutoSize = iTextFrame2.AutoSize
        .HasText = iTextFrame2.HasText
        .HorizontalAnchor = iTextFrame2.HorizontalAnchor
        .MarginBottom = iTextFrame2.MarginBottom
        .MarginLeft = iTextFrame2.MarginLeft
        .MarginRight = iTextFrame2.MarginRight
        .MarginTop = iTextFrame2.MarginTop
        '??? .NoTextRotation = iTextFrame2.NoTextRotation ' Returns or sets a value that specifies if the text on a shape is rotated if the shape itself is being rotated. Read/write.
        .Orientation = iTextFrame2.Orientation
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


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


Public Function MR_Compare(iTextFrame2 As iTextFrame2) As MR_Diff

    Dim oDiff As MR_Diff

    On Error GoTo err_

    Set oDiff = New_MR_Diff(Me, iTextFrame2)
    Set MR_Compare = oDiff

    If IsCompared(Me) Then Exit Function

    With Me
        If .AutoSize <> iTextFrame2.AutoSize Then
            Call oDiff.AddScalarProperty("AutoSize", MsoAutoSize(.AutoSize))
        End If
        'If .HasText <> iTextFrame2.HasText Then
        '    Call oDiff.AddScalarProperty("HasText", MsoTriState(.HasText))
        'End If
        If .HorizontalAnchor <> iTextFrame2.HorizontalAnchor Then
            Call oDiff.AddScalarProperty("HorizontalAnchor", MsoHorizontalAnchor(.HorizontalAnchor))
        End If
        If .MarginBottom <> iTextFrame2.MarginBottom Then
            Call oDiff.AddScalarProperty("MarginBottom", SingleToVBA(.MarginBottom))
        End If
        If .MarginLeft <> iTextFrame2.MarginLeft Then
            Call oDiff.AddScalarProperty("MarginLeft", SingleToVBA(.MarginLeft))
        End If
        If .MarginRight <> iTextFrame2.MarginRight Then
            Call oDiff.AddScalarProperty("MarginRight", SingleToVBA(.MarginRight))
        End If
        If .MarginTop <> iTextFrame2.MarginTop Then
            Call oDiff.AddScalarProperty("MarginTop", SingleToVBA(.MarginTop))
        End If
        If .Orientation <> iTextFrame2.Orientation Then
            Call oDiff.AddScalarProperty("Orientation", MsoTextOrientation(.Orientation))
        End If
        If .PathFormat <> iTextFrame2.PathFormat Then
            Call oDiff.AddScalarProperty("PathFormat", MsoPathFormat(.PathFormat))
        End If
        Call oDiff.AddDiff("TextRange", .TextRange.MR_Compare(iTextFrame2.TextRange))
        If .VerticalAnchor <> iTextFrame2.VerticalAnchor Then
            Call oDiff.AddScalarProperty("VerticalAnchor", MsoVerticalAnchor(.VerticalAnchor))
        End If
        If .WarpFormat <> iTextFrame2.WarpFormat Then
            Call oDiff.AddScalarProperty("WarpFormat", MsoWarpFormat(.WarpFormat))
        End If
        'TODO
        'If .WordArtFormat <> iTextFrame2.WordArtFormat Then
        '    Call oCode.Add("'.WordArtFormat", MsoPresetTextEffect(.WordArtFormat))
        'End If
        If .WordWrap <> iTextFrame2.WordWrap Then
            Call oDiff.AddScalarProperty("WordWrap", MsoTriState(.WordWrap))
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

Public Sub init(iTextFrame2 As TextFrame2)

    On Error GoTo err_

    belongsToDefaultShape = (TypeName(iTextFrame2.Parent.Parent) = "Presentation")

    With Me
        .AutoSize = iTextFrame2.AutoSize
        If belongsToDefaultShape = False Then
            .HasText = iTextFrame2.HasText
        End If
        .HorizontalAnchor = iTextFrame2.HorizontalAnchor
        .MarginBottom = iTextFrame2.MarginBottom
        .MarginLeft = iTextFrame2.MarginLeft
        .MarginRight = iTextFrame2.MarginRight
        .MarginTop = iTextFrame2.MarginTop
        '??? .NoTextRotation = iTextFrame2.NoTextRotation ' Returns or sets a value that specifies if the text on a shape is rotated if the shape itself is being rotated. Read/write.
        .Orientation = iTextFrame2.Orientation
        If belongsToDefaultShape = False Then
            .PathFormat = iTextFrame2.PathFormat
            Set .TextRange = New_iTextRange2(iTextFrame2.TextRange)
        Else
            Set .TextRange = New_iTextRange2()
        End If
        .VerticalAnchor = iTextFrame2.VerticalAnchor
        If belongsToDefaultShape = False Then
            .WarpFormat = iTextFrame2.WarpFormat
            .WordArtFormat = iTextFrame2.WordArtFormat
        End If
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

    Dim defaultTextFrame2 As iTextFrame2

    On Error GoTo err_

    Set defaultTextFrame2 = GetDefaultShape(Me).TextFrame2

    With Me
        .AutoSize = defaultTextFrame2.AutoSize
        .HorizontalAnchor = defaultTextFrame2.HorizontalAnchor
        .MarginBottom = defaultTextFrame2.MarginBottom
        .MarginLeft = defaultTextFrame2.MarginLeft
        .MarginRight = defaultTextFrame2.MarginRight
        .MarginTop = defaultTextFrame2.MarginTop
        .Orientation = defaultTextFrame2.Orientation
        Set .TextRange = New_iTextRange2()
        .VerticalAnchor = defaultTextFrame2.VerticalAnchor
        .WordWrap = defaultTextFrame2.WordWrap
    End With

    Exit Sub

err_:
    #If DEBUG_MODE = 1 Then
        Stop
    #Else
        err.Raise err.number 'rethrows with same source and description
    #End If

End Sub

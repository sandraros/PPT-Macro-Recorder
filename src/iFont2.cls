VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "iFont2"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Public Allcaps As MsoTriState
Public AutoRotateNumbers As MsoTriState
Public BaselineOffset As Single
Public Bold As MsoTriState
Public Caps As MsoTextCaps
Public DoubleStrikeThrough As MsoTriState
Public Equalize As MsoTriState
Public Fill As oFillFormat ' Office.FillFormat
Public Glow As iGlowFormat
Public Highlight As iColorFormat
Public Italic As MsoTriState
Public Kerning As Single
Public Name As String
Public NameAscii As String
Public NameComplexScript As String
Public NameFarEast As String
Public NameOther As String
Public Reflection As iReflectionFormat
Public Shadow As oShadowFormat
Public Size As Single
Public Smallcaps As MsoTriState
Public SoftEdgeFormat As MsoSoftEdgeType
Public Spacing As Single
Public Strike As MsoTextStrike
Public Strikethrough As MsoTriState
Public Subscript As MsoTriState
Public Superscript As MsoTriState
Public UnderlineColor As oColorFormat
Public UnderlineStyle As MsoTextUnderlineType
Public WordArtFormat As MsoPresetTextEffect


'ActiveWindow.Selection.ShapeRange.TextFrame2.TextRange.Font


Public Function MR_Compare(iFont2 As iFont2) As MR_Diff

    Dim oDiff As MR_Diff

    On Error GoTo err_

    Set oDiff = New_MR_Diff(Me, iFont2)
    Set MR_Compare = oDiff

    If IsCompared(Me) Then Exit Function

    With Me
        If .Allcaps <> iFont2.Allcaps Then
            Call oDiff.AddScalarProperty("AllCaps", MsoTriState(.Allcaps))
        End If
        If .AutoRotateNumbers <> iFont2.AutoRotateNumbers Then
            Call oDiff.AddScalarProperty("AutoRotateNumbers", MsoTriState(.AutoRotateNumbers))
        End If
        If .BaselineOffset <> iFont2.BaselineOffset Then
            Call oDiff.AddScalarProperty("BaselineOffset", SingleToVBA(.BaselineOffset))
        End If
        If .Bold <> iFont2.Bold Then
            Call oDiff.AddScalarProperty("Bold", MsoTriState(.Bold))
        End If
        If .Caps <> iFont2.Caps Then
            Call oDiff.AddScalarProperty("Caps", MsoTextCaps(.Caps))
        End If
        If .DoubleStrikeThrough <> iFont2.DoubleStrikeThrough Then
            Call oDiff.AddScalarProperty("DoubleStrikeThrough", MsoTriState(.DoubleStrikeThrough))
        End If
        If .Equalize <> iFont2.Equalize Then
            Call oDiff.AddScalarProperty("Equalize", MsoTriState(.Equalize))
        End If
        Call oDiff.AddDiff("Fill", .Fill.MR_Compare(iFont2.Fill))
        Call oDiff.AddDiff("Glow", .Glow.MR_Compare(iFont2.Glow))
        Call oDiff.AddDiff("Highlight", .Highlight.MR_Compare(iFont2.Highlight))
        If .Italic <> iFont2.Italic Then
            Call oDiff.AddScalarProperty("Italic", MsoTriState(.Italic))
        End If
        If .Kerning <> iFont2.Kerning Then
            Call oDiff.AddScalarProperty("Kerning", SingleToVBA(.Kerning))
        End If
        If .Name <> iFont2.Name Then
            Call oDiff.AddScalarProperty("Name", ToVBA(.Name))
        End If
        If .NameAscii <> iFont2.NameAscii Then
            Call oDiff.AddScalarProperty("NameAscii", ToVBA(.NameAscii))
        End If
        If .NameComplexScript <> iFont2.NameComplexScript Then
            Call oDiff.AddScalarProperty("NameComplexScript", ToVBA(.NameComplexScript))
        End If
        If .NameFarEast <> iFont2.NameFarEast Then
            Call oDiff.AddScalarProperty("NameFarEast", ToVBA(.NameFarEast))
        End If
        If .NameOther <> iFont2.NameOther Then
            Call oDiff.AddScalarProperty("NameOther", ToVBA(.NameOther))
        End If
        Call oDiff.AddDiff("Reflection", .Reflection.MR_Compare(iFont2.Reflection))
        Call oDiff.AddDiff("Shadow", .Shadow.MR_Compare(iFont2.Shadow))
        If .Size <> iFont2.Size Then
            Call oDiff.AddScalarProperty("Size", SingleToVBA(.Size))
        End If
        If .Smallcaps <> iFont2.Smallcaps Then
            Call oDiff.AddScalarProperty("Smallcaps", MsoTriState(.Smallcaps))
        End If
        If .SoftEdgeFormat <> iFont2.SoftEdgeFormat Then
            Call oDiff.AddScalarProperty("SoftEdgeFormat", MsoSoftEdgeType(.SoftEdgeFormat))
        End If
        If .Spacing <> iFont2.Spacing Then
            Call oDiff.AddScalarProperty("Spacing", SingleToVBA(.Spacing))
        End If
        If .Strike <> iFont2.Strike Then
            Call oDiff.AddScalarProperty("Strike", MsoTextStrike(.Strike))
        End If
        If .Strikethrough <> iFont2.Strikethrough Then
            Call oDiff.AddScalarProperty("Strikethrough", MsoTriState(.Strikethrough))
        End If
        If .Subscript <> iFont2.Subscript Then
            Call oDiff.AddScalarProperty("Subscript", MsoTriState(.Subscript))
        End If
        If .Superscript <> iFont2.Superscript Then
            Call oDiff.AddScalarProperty("Superscript", MsoTriState(.Superscript))
        End If
        Call oDiff.AddDiff("UnderlineColor", .UnderlineColor.MR_Compare(iFont2.UnderlineColor))
        If .UnderlineStyle <> iFont2.UnderlineStyle Then
            Call oDiff.AddScalarProperty("UnderlineStyle", MsoTextUnderlineType(.UnderlineStyle))
        End If
        If .WordArtFormat <> iFont2.WordArtFormat Then
            Call oDiff.AddScalarProperty("WordArtFormat", MsoPresetTextEffect(.WordArtFormat))
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

Public Sub init(ByVal iFont2 As Office.Font2)

    On Error GoTo err_

    With Me
        .Allcaps = iFont2.Allcaps
        .AutoRotateNumbers = iFont2.AutoRotateNumbers
        .BaselineOffset = iFont2.BaselineOffset
        .Bold = iFont2.Bold
        .Caps = iFont2.Caps
        .DoubleStrikeThrough = iFont2.DoubleStrikeThrough
        .Equalize = iFont2.Equalize
        Set .Fill = New_oFillFormat(iFont2.Fill)
        Set .Glow = New_iGlowFormat(iFont2.Glow)
        If iFont2.Highlight.Type <> msoColorTypeMixed Then
            Set .Highlight = New_iColorFormat(iFont2.Highlight)
        Else
            Set .Highlight = New iColorFormat
        End If
        .Italic = iFont2.Italic
        .Kerning = iFont2.Kerning
        .Name = iFont2.Name
        .NameAscii = iFont2.NameAscii
        .NameComplexScript = iFont2.NameComplexScript
        .NameFarEast = iFont2.NameFarEast
        .NameOther = iFont2.NameOther
        Set .Reflection = New_iReflectionFormat(iFont2.Reflection)
        Set .Shadow = New_oShadowFormat(iFont2.Shadow)
        .Size = iFont2.Size
        .Smallcaps = iFont2.Smallcaps
        .SoftEdgeFormat = iFont2.SoftEdgeFormat
        .Spacing = iFont2.Spacing
        .Strike = iFont2.Strike
        .Strikethrough = iFont2.Strikethrough
        .Subscript = iFont2.Subscript
        .Superscript = iFont2.Superscript
        Set .UnderlineColor = New_oColorFormat(iFont2.UnderlineColor)
        .UnderlineStyle = iFont2.UnderlineStyle
        If iFont2.WordArtFormat <> msoTextEffectMixed Then
            .WordArtFormat = iFont2.WordArtFormat
        End If
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
    
    Dim defaultFont2 As iFont2

    On Error GoTo err_

    If comparisonRunning = False Then

        With Me
            Set .Fill = New_oFillFormat()
            Set .Glow = New_iGlowFormat()
            Set .Highlight = New_iColorFormat()
            Set .Reflection = New_iReflectionFormat()
            Set .Shadow = New_oShadowFormat()
            Set .UnderlineColor = New_oColorFormat()
        End With

    Else

        Set defaultFont2 = GetDefaultShape(Me).TextFrame2.TextRange.Font

        With Me
            Set .Fill = defaultFont2.Fill
            Set .Glow = defaultFont2.Glow
            Set .Highlight = New_iColorFormat()
            Set .Reflection = New_iReflectionFormat()
            Set .Shadow = New_oShadowFormat()
            Set .UnderlineColor = New_oColorFormat()
        End With

    End If

    Exit Sub

err_:
    #If DEBUG_MODE = 1 Then
        Stop
    #Else
        err.Raise err.number 'rethrows with same source and description
    #End If

End Sub


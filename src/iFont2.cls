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


Public Function Compare(ObjectName As String, iFont2 As iFont2) As UDiff

    Dim oCode As UDiff

    On Error GoTo err_

    Set oCode = New_UDiff(ObjectName, Me, iFont2)
    Set Compare = oCode

    If IsCompared(Me) Then Exit Function

    With Me
        If .Allcaps <> iFont2.Allcaps Then
            Call oCode.AddScalarProperty("AllCaps", MsoTriState(.Allcaps))
        End If
        If .AutoRotateNumbers <> iFont2.AutoRotateNumbers Then
            Call oCode.AddScalarProperty("AutoRotateNumbers", MsoTriState(.AutoRotateNumbers))
        End If
        If .BaselineOffset <> iFont2.BaselineOffset Then
            Call oCode.AddScalarProperty("BaselineOffset", SingleToVBA(.BaselineOffset))
        End If
        If .Bold <> iFont2.Bold Then
            Call oCode.AddScalarProperty("Bold", MsoTriState(.Bold))
        End If
        If .Caps <> iFont2.Caps Then
            Call oCode.AddScalarProperty("Caps", MsoTextCaps(.Caps))
        End If
        If .DoubleStrikeThrough <> iFont2.DoubleStrikeThrough Then
            Call oCode.AddScalarProperty("DoubleStrikeThrough", MsoTriState(.DoubleStrikeThrough))
        End If
        If .Equalize <> iFont2.Equalize Then
            Call oCode.AddScalarProperty("Equalize", MsoTriState(.Equalize))
        End If
        Call oCode.AddCode(.Fill.Compare("Fill", iFont2.Fill))
        Call oCode.AddCode(.Glow.Compare("Glow", iFont2.Glow))
        Call oCode.AddCode(.Highlight.Compare("Highlight", iFont2.Highlight))
        If .Italic <> iFont2.Italic Then
            Call oCode.AddScalarProperty("Italic", MsoTriState(.Italic))
        End If
        If .Kerning <> iFont2.Kerning Then
            Call oCode.AddScalarProperty("Kerning", SingleToVBA(.Kerning))
        End If
        If .Name <> iFont2.Name Then
            Call oCode.AddScalarProperty("Name", ToVBA(.Name))
        End If
        If .NameAscii <> iFont2.NameAscii Then
            Call oCode.AddScalarProperty("NameAscii", ToVBA(.NameAscii))
        End If
        If .NameComplexScript <> iFont2.NameComplexScript Then
            Call oCode.AddScalarProperty("NameComplexScript", ToVBA(.NameComplexScript))
        End If
        If .NameFarEast <> iFont2.NameFarEast Then
            Call oCode.AddScalarProperty("NameFarEast", ToVBA(.NameFarEast))
        End If
        If .NameOther <> iFont2.NameOther Then
            Call oCode.AddScalarProperty("NameOther", ToVBA(.NameOther))
        End If
        Call oCode.AddCode(.Reflection.Compare("Reflection", iFont2.Reflection))
        Call oCode.AddCode(.Shadow.Compare("Shadow", iFont2.Shadow))
        If .Size <> iFont2.Size Then
            Call oCode.AddScalarProperty("Size", SingleToVBA(.Size))
        End If
        If .Smallcaps <> iFont2.Smallcaps Then
            Call oCode.AddScalarProperty("Smallcaps", MsoTriState(.Smallcaps))
        End If
        If .SoftEdgeFormat <> iFont2.SoftEdgeFormat Then
            Call oCode.AddScalarProperty("SoftEdgeFormat", MsoSoftEdgeType(.SoftEdgeFormat))
        End If
        If .Spacing <> iFont2.Spacing Then
            Call oCode.AddScalarProperty("Spacing", SingleToVBA(.Spacing))
        End If
        If .Strike <> iFont2.Strike Then
            Call oCode.AddScalarProperty("Strike", MsoTextStrike(.Strike))
        End If
        If .Strikethrough <> iFont2.Strikethrough Then
            Call oCode.AddScalarProperty("Strikethrough", MsoTriState(.Strikethrough))
        End If
        If .Subscript <> iFont2.Subscript Then
            Call oCode.AddScalarProperty("Subscript", MsoTriState(.Subscript))
        End If
        If .Superscript <> iFont2.Superscript Then
            Call oCode.AddScalarProperty("Superscript", MsoTriState(.Superscript))
        End If
        Call oCode.AddCode(.UnderlineColor.Compare("UnderlineColor", iFont2.UnderlineColor))
        If .UnderlineStyle <> iFont2.UnderlineStyle Then
            Call oCode.AddScalarProperty("UnderlineStyle", MsoTextUnderlineType(.UnderlineStyle))
        End If
        If .WordArtFormat <> iFont2.WordArtFormat Then
            Call oCode.AddScalarProperty("WordArtFormat", MsoPresetTextEffect(.WordArtFormat))
        End If
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

    If comparisonRunning = False Then

        With Me
            Set .Fill = New oFillFormat
            Call .Fill.DefaultValues
            Set .Glow = New iGlowFormat
            Call .Glow.DefaultValues
            Set .Highlight = New iColorFormat
            Call .Highlight.DefaultValues
            Set .Reflection = New iReflectionFormat
            Call .Reflection.DefaultValues
            Set .Shadow = New oShadowFormat
            Call .Shadow.DefaultValues
            Set .UnderlineColor = New oColorFormat
            Call .UnderlineColor.DefaultValues
        End With

    Else

        Set defaultFont2 = defaultShape.TextFrame2.TextRange.Font

        With Me
            Set .Fill = defaultFont2.Fill
            Set .Glow = defaultFont2.Glow
            Set .Highlight = New iColorFormat
            Call .Highlight.DefaultValues
            Set .Reflection = New iReflectionFormat
            Call .Reflection.DefaultValues
            Set .Shadow = New oShadowFormat
            Call .Shadow.DefaultValues
            Set .UnderlineColor = New oColorFormat
            Call .UnderlineColor.DefaultValues
        End With

    End If

End Sub



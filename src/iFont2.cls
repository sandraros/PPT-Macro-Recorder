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


Public Function compare(propertyName As String, iFont2 As iFont2) As Code

    Dim objCode As Code

    On Error GoTo err_

    Set objCode = New Code
    Set compare = objCode

    If IsCompared(Me) Then Exit Function

    With Me
        If .Allcaps <> iFont2.Allcaps Then
            Call objCode.Add(".AllCaps = " & MsoTriState(.Allcaps))
        End If
        If .AutoRotateNumbers <> iFont2.AutoRotateNumbers Then
            Call objCode.Add(".AutoRotateNumbers = " & MsoTriState(.AutoRotateNumbers))
        End If
        If .BaselineOffset <> iFont2.BaselineOffset Then
            Call objCode.Add(".BaselineOffset = " & SingleToVBA(.BaselineOffset))
        End If
        If .Bold <> iFont2.Bold Then
            Call objCode.Add(".Bold = " & MsoTriState(.Bold))
        End If
        If .Caps <> iFont2.Caps Then
            Call objCode.Add(".Caps = " & MsoTextCaps(.Caps))
        End If
        If .DoubleStrikeThrough <> iFont2.DoubleStrikeThrough Then
            Call objCode.Add(".DoubleStrikeThrough = " & MsoTriState(.DoubleStrikeThrough))
        End If
        If .Equalize <> iFont2.Equalize Then
            Call objCode.Add(".Equalize = " & MsoTriState(.Equalize))
        End If
        Call objCode.AddCode(.Fill.compare(".Fill", iFont2.Fill))
        Call objCode.AddCode(.Glow.compare(".Glow", iFont2.Glow))
        Call objCode.AddCode(.Highlight.compare(".Highlight", iFont2.Highlight))
        If .Italic <> iFont2.Italic Then
            Call objCode.Add(".Italic = " & MsoTriState(.Italic))
        End If
        If .Kerning <> iFont2.Kerning Then
            Call objCode.Add(".Kerning = " & SingleToVBA(.Kerning))
        End If
        If .Name <> iFont2.Name Then
            Call objCode.Add(".Name = " & ToVBA(.Name))
        End If
        If .NameAscii <> iFont2.NameAscii Then
            Call objCode.Add(".NameAscii = " & ToVBA(.NameAscii))
        End If
        If .NameComplexScript <> iFont2.NameComplexScript Then
            Call objCode.Add(".NameComplexScript = " & ToVBA(.NameComplexScript))
        End If
        If .NameFarEast <> iFont2.NameFarEast Then
            Call objCode.Add(".NameFarEast = " & ToVBA(.NameFarEast))
        End If
        If .NameOther <> iFont2.NameOther Then
            Call objCode.Add(".NameOther = " & ToVBA(.NameOther))
        End If
        Call objCode.AddCode(.Reflection.compare(".Reflection", iFont2.Reflection))
        Call objCode.AddCode(.Shadow.compare(".Shadow", iFont2.Shadow))
        If .Size <> iFont2.Size Then
            Call objCode.Add(".Size = " & SingleToVBA(.Size))
        End If
        If .Smallcaps <> iFont2.Smallcaps Then
            Call objCode.Add(".Smallcaps = " & MsoTriState(.Smallcaps))
        End If
        If .SoftEdgeFormat <> iFont2.SoftEdgeFormat Then
            Call objCode.Add(".SoftEdgeFormat = " & MsoSoftEdgeType(.SoftEdgeFormat))
        End If
        If .Spacing <> iFont2.Spacing Then
            Call objCode.Add(".Spacing = " & SingleToVBA(.Spacing))
        End If
        If .Strike <> iFont2.Strike Then
            Call objCode.Add(".Strike = " & MsoTextStrike(.Strike))
        End If
        If .Strikethrough <> iFont2.Strikethrough Then
            Call objCode.Add(".Strikethrough = " & MsoTriState(.Strikethrough))
        End If
        If .Subscript <> iFont2.Subscript Then
            Call objCode.Add(".Subscript = " & MsoTriState(.Subscript))
        End If
        If .Superscript <> iFont2.Superscript Then
            Call objCode.Add(".Superscript = " & MsoTriState(.Superscript))
        End If
        Call objCode.AddCode(.UnderlineColor.compare(".UnderlineColor", iFont2.UnderlineColor))
        If .UnderlineStyle <> iFont2.UnderlineStyle Then
            Call objCode.Add(".UnderlineStyle= " & MsoTextUnderlineType(.UnderlineStyle))
        End If
        If .WordArtFormat <> iFont2.WordArtFormat Then
            Call objCode.Add(".WordArtFormat = " & MsoPresetTextEffect(.WordArtFormat))
        End If
    End With

    Call objCode.Wrap(propertyName)


    Set compare = objCode

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
End Sub



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


Public Function compare(propertyName As String, iFont2 As iFont2, indent As Integer) As String

    On Error GoTo err_

    If IsCompared(Me) Then Exit Function

    code = ""

    With Me
        If .Allcaps <> iFont2.Allcaps Then
            code = code & Space(indent) & ".AllCaps = " & MsoTriState(.Allcaps) & Chr(13)
        End If
        If .AutoRotateNumbers <> iFont2.AutoRotateNumbers Then
            code = code & Space(indent) & ".AutoRotateNumbers = " & MsoTriState(.AutoRotateNumbers) & Chr(13)
        End If
        If .BaselineOffset <> iFont2.BaselineOffset Then
            code = code & Space(indent) & ".BaselineOffset = " & SingleToVBA(.BaselineOffset) & Chr(13)
        End If
        If .Bold <> iFont2.Bold Then
            code = code & Space(indent) & ".Bold = " & MsoTriState(.Bold) & Chr(13)
        End If
        If .Caps <> iFont2.Caps Then
            code = code & Space(indent) & ".Caps = " & MsoTextCaps(.Caps) & Chr(13)
        End If
        If .DoubleStrikeThrough <> iFont2.DoubleStrikeThrough Then
            code = code & Space(indent) & ".DoubleStrikeThrough = " & MsoTriState(.DoubleStrikeThrough) & Chr(13)
        End If
        If .Equalize <> iFont2.Equalize Then
            code = code & Space(indent) & ".Equalize = " & MsoTriState(.Equalize) & Chr(13)
        End If
        code = code & .Fill.compare(".Fill", iFont2.Fill, indent + 4)
        code = code & .Glow.compare(".Glow", iFont2.Glow, indent + 4)
        code = code & .Highlight.compare(".Highlight", iFont2.Highlight, indent + 4)
        If .Italic <> iFont2.Italic Then
            code = code & Space(indent) & ".Italic = " & MsoTriState(.Italic) & Chr(13)
        End If
        If .Kerning <> iFont2.Kerning Then
            code = code & Space(indent) & ".Kerning = " & SingleToVBA(.Kerning) & Chr(13)
        End If
        If .Name <> iFont2.Name Then
            code = code & Space(indent) & ".Name = " & ToVBA(.Name) & Chr(13)
        End If
        If .NameAscii <> iFont2.NameAscii Then
            code = code & Space(indent) & ".NameAscii = " & ToVBA(.NameAscii) & Chr(13)
        End If
        If .NameComplexScript <> iFont2.NameComplexScript Then
            code = code & Space(indent) & ".NameComplexScript = " & ToVBA(.NameComplexScript) & Chr(13)
        End If
        If .NameFarEast <> iFont2.NameFarEast Then
            code = code & Space(indent) & ".NameFarEast = " & ToVBA(.NameFarEast) & Chr(13)
        End If
        If .NameOther <> iFont2.NameOther Then
            code = code & Space(indent) & ".NameOther = " & ToVBA(.NameOther) & Chr(13)
        End If
        code = code & .Reflection.compare(".Reflection", iFont2.Reflection, indent + 4)
        code = code & .Shadow.compare(".Shadow", iFont2.Shadow, indent + 4)
        If .Size <> iFont2.Size Then
            code = code & Space(indent) & ".Size = " & SingleToVBA(.Size) & Chr(13)
        End If
        If .Smallcaps <> iFont2.Smallcaps Then
            code = code & Space(indent) & ".Smallcaps = " & MsoTriState(.Smallcaps) & Chr(13)
        End If
        If .SoftEdgeFormat <> iFont2.SoftEdgeFormat Then
            code = code & Space(indent) & ".SoftEdgeFormat = " & MsoSoftEdgeType(.SoftEdgeFormat) & Chr(13)
        End If
        If .Spacing <> iFont2.Spacing Then
            code = code & Space(indent) & ".Spacing = " & SingleToVBA(.Spacing) & Chr(13)
        End If
        If .Strike <> iFont2.Strike Then
            code = code & Space(indent) & ".Strike = " & MsoTextStrike(.Strike) & Chr(13)
        End If
        If .Strikethrough <> iFont2.Strikethrough Then
            code = code & Space(indent) & ".Strikethrough = " & MsoTriState(.Strikethrough) & Chr(13)
        End If
        If .Subscript <> iFont2.Subscript Then
            code = code & Space(indent) & ".Subscript = " & MsoTriState(.Subscript) & Chr(13)
        End If
        If .Superscript <> iFont2.Superscript Then
            code = code & Space(indent) & ".Superscript = " & MsoTriState(.Superscript) & Chr(13)
        End If
        code = code & .UnderlineColor.compare(".UnderlineColor", iFont2.UnderlineColor, indent + 4)
        If .UnderlineStyle <> iFont2.UnderlineStyle Then
            code = code & Space(indent) & ".UnderlineStyle= " & MsoTextUnderlineType(.UnderlineStyle) & Chr(13)
        End If
        If .WordArtFormat <> iFont2.WordArtFormat Then
            code = code & Space(indent) & ".WordArtFormat = " & MsoPresetTextEffect(.WordArtFormat) & Chr(13)
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



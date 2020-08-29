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
Public Fill As iFillFormat
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
Public Shadow As iShadowFormat
Public Size As Single
Public Smallcaps As MsoTriState
Public SoftEdgeFormat As MsoSoftEdgeType
Public Spacing As Single
Public Strike As MsoTextStrike
Public Strikethrough As MsoTriState
Public Subscript As MsoTriState
Public Superscript As MsoTriState
Public UnderlineColor As iColorFormat2
Public UnderlineStyle As MsoTextUnderlineType
Public WordArtFormat As MsoPresetTextEffect

Public Sub class_initialize()

    With Me
        Set .Fill = New iFillFormat
        Set .Glow = New iGlowFormat
        Set .Highlight = New iColorFormat
        Set .Reflection = New iReflectionFormat
        Set .Shadow = New iShadowFormat
        Set .UnderlineColor = New iColorFormat2
    End With

End Sub

Public Function compare(ByVal iFont2 As iFont2, indent As Integer) As String

    On Error GoTo err_

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
        code = code & ObjectToVBA("Fill", .Fill.compare(iFont2.Fill, indent))
        code = code & ObjectToVBA("Glow", .Glow.compare(iFont2.Glow, indent))
        code = code & ObjectToVBA("Highlight", .Highlight.compare(iFont2.Highlight, indent))
        If .Italic <> iFont2.Italic Then
            code = code & Space(indent) & ".Italic = " & MsoTriState(.Italic) & Chr(13)
        End If
        If .Kerning <> iFont2.Kerning Then
            code = code & Space(indent) & ".Kerning = " & SingleToVBA(.Kerning) & Chr(13)
        End If
        If .Name <> iFont2.Name Then
            code = code & Space(indent) & ".Name = " & .Name & Chr(13)
        End If
        If .NameAscii <> iFont2.NameAscii Then
            code = code & Space(indent) & ".NameAscii = " & .NameAscii & Chr(13)
        End If
        If .NameComplexScript <> iFont2.NameComplexScript Then
            code = code & Space(indent) & ".NameComplexScript = " & .NameComplexScript & Chr(13)
        End If
        If .NameFarEast <> iFont2.NameFarEast Then
            code = code & Space(indent) & ".NameFarEast = " & .NameFarEast & Chr(13)
        End If
        If .NameOther <> iFont2.NameOther Then
            code = code & Space(indent) & ".NameOther = " & .NameOther & Chr(13)
        End If
        code = code & ObjectToVBA("Reflection", .Reflection.compare(iFont2.Reflection, indent))
        code = code & ObjectToVBA("Shadow", .Shadow.compare(iFont2.Shadow, indent))
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
        code = code & ObjectToVBA("UnderlineColor", .UnderlineColor.compare(iFont2.UnderlineColor, indent))
        If .UnderlineStyle <> iFont2.UnderlineStyle Then
            code = code & Space(indent) & ".UnderlineStyle= " & MsoTextUnderlineType(.UnderlineStyle) & Chr(13)
        End If
        If .WordArtFormat <> iFont2.WordArtFormat Then
            code = code & Space(indent) & ".WordArtFormat = " & MsoPresetTextEffect(.WordArtFormat) & Chr(13)
        End If
    End With

    compare = code

    Exit Function

err_:
    #If DEBUG_MODE = 1 Then
        Stop
    #Else
        err.Raise err.number 'rethrows with same source and description
    #End If

End Function

Public Sub Init(ByVal iFont2 As Office.Font2)

    On Error GoTo err_

    With Me
        .Allcaps = iFont2.Allcaps
        .AutoRotateNumbers = iFont2.AutoRotateNumbers
        .BaselineOffset = iFont2.BaselineOffset
        .Bold = iFont2.Bold
        .Caps = iFont2.Caps
        .DoubleStrikeThrough = iFont2.DoubleStrikeThrough
        .Equalize = iFont2.Equalize
        If iFont2.Fill.ForeColor.Type <> msoColorTypeMixed Then
            Call .Fill.Init(iFont2.Fill)
        End If
        If iFont2.Glow.Color.Type <> msoColorTypeMixed Then
            Call .Glow.Init(iFont2.Glow)
        End If
        If iFont2.Highlight.Type <> msoColorTypeMixed Then
            Call .Highlight.Init(iFont2.Highlight)
        End If
        .Italic = iFont2.Italic
        .Kerning = iFont2.Kerning
        .Name = iFont2.Name
        .NameAscii = iFont2.NameAscii
        .NameComplexScript = iFont2.NameComplexScript
        .NameFarEast = iFont2.NameFarEast
        .NameOther = iFont2.NameOther
        Call .Reflection.Init(iFont2.Reflection)
        Call .Shadow.Init(iFont2.Shadow)
        .Size = iFont2.Size
        .Smallcaps = iFont2.Smallcaps
        .SoftEdgeFormat = iFont2.SoftEdgeFormat
        .Spacing = iFont2.Spacing
        .Strike = iFont2.Strike
        .Strikethrough = iFont2.Strikethrough
        .Subscript = iFont2.Subscript
        If iFont2.UnderlineColor.Type <> msoColorTypeMixed Then
            Call .UnderlineColor.Init(iFont2.UnderlineColor)
        End If
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


VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "iFont"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Public AutoRotateNumbers As MsoTriState
Public BaselineOffset As Single
Public Bold As MsoTriState
Public Color As iColorFormat
Public Embeddable As MsoTriState
Public Embedded As MsoTriState
Public Emboss As MsoTriState
Public Italic As MsoTriState
Public Name As String
Public NameAscii As String
Public NameComplexScript As String
Public NameFarEast As String
Public NameOther As String
Public Parent As Object
Public Shadow As MsoTriState
Public Size As Single
Public Subscript As MsoTriState
Public Superscript As MsoTriState
Public Underline As MsoTriState


'ActiveWindow.Selection.ShapeRange.TextFrame.TextRange.Font.AutoRotateNumbers


Public Function Compare(ObjectName As String, iFont As iFont) As UDiff

    Dim oCode As UDiff

    On Error GoTo err_

    Set oCode = New_UDiff(ObjectName, Me, iFont)
    Set Compare = oCode

    If IsCompared(Me) Then Exit Function

    With Me
        If .AutoRotateNumbers <> iFont.AutoRotateNumbers Then
            Call oCode.AddScalarProperty("AutoRotateNumbers", MsoTriState(.AutoRotateNumbers))
        End If
        If .BaselineOffset <> iFont.BaselineOffset Then
            Call oCode.AddScalarProperty("BaselineOffset", SingleToVBA(.BaselineOffset))
        End If
        If .Bold <> iFont.Bold Then
            Call oCode.AddScalarProperty("Bold", MsoTriState(.Bold))
        End If
        Call oCode.AddCode(.Color.Compare("Color", iFont.Color))
        If .Embeddable <> iFont.Embeddable Then
            Call oCode.AddScalarProperty("Embeddable", MsoTriState(.Embeddable))
        End If
        If .Embedded <> iFont.Embedded Then
            Call oCode.AddScalarProperty("Embedded", MsoTriState(.Embedded))
        End If
        If .Emboss <> iFont.Emboss Then
            Call oCode.AddScalarProperty("Emboss", MsoTriState(.Emboss))
        End If
        If .Italic <> iFont.Italic Then
            Call oCode.AddScalarProperty("Italic", MsoTriState(.Italic))
        End If
        If .Name <> iFont.Name Then
            Call oCode.AddScalarProperty("Name", .Name)
        End If
        If .NameAscii <> iFont.NameAscii Then
            Call oCode.AddScalarProperty("NameAscii", .NameAscii)
        End If
        If .NameComplexScript <> iFont.NameComplexScript Then
            Call oCode.AddScalarProperty("NameComplexScript", .NameComplexScript)
        End If
        If .NameFarEast <> iFont.NameFarEast Then
            Call oCode.AddScalarProperty("NameFarEast", .NameFarEast)
        End If
        If .NameOther <> iFont.NameOther Then
            Call oCode.AddScalarProperty("NameOther", .NameOther)
        End If
        'If Not .Parent Is iFont.Parent Then
        '    Call oCode.Add("'TODO Parent")
        'End If
        If .Shadow <> iFont.Shadow Then
            Call oCode.AddScalarProperty("Shadow", .Shadow)
        End If
        If .Size <> iFont.Size Then
            Call oCode.AddScalarProperty("Size", SingleToVBA(.Size))
        End If
        If .Subscript <> iFont.Subscript Then
            Call oCode.AddScalarProperty("Subscript", MsoTriState(.Subscript))
        End If
        If .Superscript <> iFont.Superscript Then
            Call oCode.AddScalarProperty("Superscript", MsoTriState(.Superscript))
        End If
        If .Underline <> iFont.Underline Then
            Call oCode.AddScalarProperty("Underline", MsoTriState(.Underline))
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

Public Sub init(ByVal iFont As Font)

    On Error GoTo err_

    With Me
        .AutoRotateNumbers = iFont.AutoRotateNumbers
        .BaselineOffset = iFont.BaselineOffset
        Set .Color = New_iColorFormat(iFont.Color)
        .Embeddable = iFont.Embeddable
        .Embedded = iFont.Embedded
        .Emboss = iFont.Emboss
        .Italic = iFont.Italic
        .Name = iFont.Name
        .NameAscii = iFont.NameAscii
        .NameComplexScript = iFont.NameComplexScript
        .NameFarEast = iFont.NameFarEast
        .NameOther = iFont.NameOther
        Set .Parent = iFont.Parent
        .Shadow = iFont.Shadow
        .Size = iFont.Size
        .Subscript = iFont.Subscript
        .Superscript = iFont.Superscript
        .Underline = iFont.Underline
    End With

    Exit Sub

err_:
    #If DEBUG_MODE = 1 Then
        Stop
    #Else
        err.Raise err.number 'rethrows with same source and description
    #End If

End Sub



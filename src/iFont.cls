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


Public Function compare(propertyName As String, iFont As iFont) As Code

    Dim objCode As Code

    On Error GoTo err_

    Set objCode = New Code
    Set compare = objCode

    If IsCompared(Me) Then Exit Function

    With Me
        If .AutoRotateNumbers <> iFont.AutoRotateNumbers Then
            Call objCode.Add(".AutoRotateNumbers = " & MsoTriState(.AutoRotateNumbers))
        End If
        If .BaselineOffset <> iFont.BaselineOffset Then
            Call objCode.Add(".BaselineOffset = " & SingleToVBA(.BaselineOffset))
        End If
        If .Bold <> iFont.Bold Then
            Call objCode.Add(".Bold = " & MsoTriState(.Bold))
        End If
        Call objCode.AddCode(.Color.compare(".Color", iFont.Color))
        If .Embeddable <> iFont.Embeddable Then
            Call objCode.Add(".Embeddable = " & MsoTriState(.Embeddable))
        End If
        If .Embedded <> iFont.Embedded Then
            Call objCode.Add(".Embedded = " & MsoTriState(.Embedded))
        End If
        If .Emboss <> iFont.Emboss Then
            Call objCode.Add(".Emboss = " & MsoTriState(.Emboss))
        End If
        If .Italic <> iFont.Italic Then
            Call objCode.Add(".Italic = " & MsoTriState(.Italic))
        End If
        If .Name <> iFont.Name Then
            Call objCode.Add(".Name = " & .Name)
        End If
        If .NameAscii <> iFont.NameAscii Then
            Call objCode.Add(".NameAscii = " & .NameAscii)
        End If
        If .NameComplexScript <> iFont.NameComplexScript Then
            Call objCode.Add(".NameComplexScript = " & .NameComplexScript)
        End If
        If .NameFarEast <> iFont.NameFarEast Then
            Call objCode.Add(".NameFarEast = " & .NameFarEast)
        End If
        If .NameOther <> iFont.NameOther Then
            Call objCode.Add(".NameOther = " & .NameOther)
        End If
        'If Not .Parent Is iFont.Parent Then
        '    Call objCode.Add("'TODO Parent")
        'End If
        If .Shadow <> iFont.Shadow Then
            Call objCode.Add(".Shadow = " & .Shadow)
        End If
        If .Size <> iFont.Size Then
            Call objCode.Add(".Size = " & SingleToVBA(.Size))
        End If
        If .Subscript <> iFont.Subscript Then
            Call objCode.Add(".Subscript = " & MsoTriState(.Subscript))
        End If
        If .Superscript <> iFont.Superscript Then
            Call objCode.Add(".Superscript = " & MsoTriState(.Superscript))
        End If
        If .Underline <> iFont.Underline Then
            Call objCode.Add(".Underline = " & MsoTriState(.Underline))
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



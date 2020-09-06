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


Public Function compare(propertyName As String, iFont As iFont, indent As Integer) As String

    On Error GoTo err_

    If IsCompared(Me) Then Exit Function

    code = ""

    With Me
        If .AutoRotateNumbers <> iFont.AutoRotateNumbers Then
            code = code & Space(indent) & ".AutoRotateNumbers = " & MsoTriState(.AutoRotateNumbers) & Chr(13)
        End If
        If .BaselineOffset <> iFont.BaselineOffset Then
            code = code & Space(indent) & ".BaselineOffset = " & SingleToVBA(.BaselineOffset) & Chr(13)
        End If
        If .Bold <> iFont.Bold Then
            code = code & Space(indent) & ".Bold = " & MsoTriState(.Bold) & Chr(13)
        End If
        code = code & .Color.compare(".Color", iFont.Color, indent + 4)
        If .Embeddable <> iFont.Embeddable Then
            code = code & Space(indent) & ".Embeddable = " & MsoTriState(.Embeddable) & Chr(13)
        End If
        If .Embedded <> iFont.Embedded Then
            code = code & Space(indent) & ".Embedded = " & MsoTriState(.Embedded) & Chr(13)
        End If
        If .Emboss <> iFont.Emboss Then
            code = code & Space(indent) & ".Emboss = " & MsoTriState(.Emboss) & Chr(13)
        End If
        If .Italic <> iFont.Italic Then
            code = code & Space(indent) & ".Italic = " & MsoTriState(.Italic) & Chr(13)
        End If
        If .Name <> iFont.Name Then
            code = code & Space(indent) & ".Name = " & .Name & Chr(13)
        End If
        If .NameAscii <> iFont.NameAscii Then
            code = code & Space(indent) & ".NameAscii = " & .NameAscii & Chr(13)
        End If
        If .NameComplexScript <> iFont.NameComplexScript Then
            code = code & Space(indent) & ".NameComplexScript = " & .NameComplexScript & Chr(13)
        End If
        If .NameFarEast <> iFont.NameFarEast Then
            code = code & Space(indent) & ".NameFarEast = " & .NameFarEast & Chr(13)
        End If
        If .NameOther <> iFont.NameOther Then
            code = code & Space(indent) & ".NameOther = " & .NameOther & Chr(13)
        End If
        If Not .Parent Is iFont.Parent Then
            code = code & "'TODO Parent"
        End If
        code = code & .Shadow.compare(".Shadow", iFont.Shadow, indent + 4)
        If .Size <> iFont.Size Then
            code = code & Space(indent) & ".Size = " & SingleToVBA(.Size) & Chr(13)
        End If
        If .Subscript <> iFont.Subscript Then
            code = code & Space(indent) & ".Subscript = " & MsoTriState(.Subscript) & Chr(13)
        End If
        If .Superscript <> iFont.Superscript Then
            code = code & Space(indent) & ".Superscript = " & MsoTriState(.Superscript) & Chr(13)
        End If
        If .Underline <> iFont.Underline Then
            code = code & Space(indent) & ".Underline = " & MsoTriState(.Underline) & Chr(13)
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



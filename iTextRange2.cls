VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "iTextRange2"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Public Application As Object
Public BoundHeight As Single
Public BoundLeft As Single
Public BoundTop As Single
Public BoundWidth As Single
Public Count As Long
Public Creator As Long
Public Font As iFont2
Public LanguageID As MsoLanguageID
Public length As Long
'Public ParagraphFormat As iParagraphFormat2
Public Parent As Object
Public Start As Long
Public Text As String


'ActiveWindow.Selection.ShapeRange.TextFrame2.TextRange.Start


Public Function compare(propertyName As String, iTextRange2 As iTextRange2, indent As Integer) As String

    On Error GoTo err_

    If IsCompared(Me) Then Exit Function

    code = ""

    With Me
        'If .Application <> iTextRange2.Application Then
        '    code = code & ".Application = " & iTextRange2.Application & Chr(13)
        'End If
        'If .BoundHeight <> iTextRange2.BoundHeight Then
        '    code = code & Space(indent) & ".BoundHeight = " & SingleToVBA(.BoundHeight) & Chr(13)
        'End If
        'If .BoundLeft <> iTextRange2.BoundLeft Then
        '    code = code & Space(indent) & ".BoundLeft = " & SingleToVBA(.BoundLeft) & Chr(13)
        'End If
        'If .BoundTop <> iTextRange2.BoundTop Then
        '    code = code & Space(indent) & ".BoundTop = " & SingleToVBA(.BoundTop) & Chr(13)
        'End If
        'If .BoundWidth <> iTextRange2.BoundWidth Then
        '    code = code & Space(indent) & ".BoundWidth = " & SingleToVBA(.BoundWidth) & Chr(13)
        'End If
        'If .Count <> iTextRange2.Count Then
        '    code = code & Space(indent) & ".Count = " & LongToVBA(.Count) & Chr(13)
        'End If
        'If .Creator <> iTextRange2.Creator Then
        '    code = code & Space(indent) & ".Creator = " & LongToVBA(.Creator) & Chr(13)
        'End If
        code = code & .Font.compare(".Font", iTextRange2.Font, indent + 4)
        If .LanguageID <> iTextRange2.LanguageID Then
            code = code & Space(indent) & ".LanguageID = " & MsoLanguageID(.LanguageID) & Chr(13)
        End If
        'If .length <> iTextRange2.length Then
        '    code = code & Space(indent) & ".length = " & LongToVBA(.length) & Chr(13)
        'End If
        'If Not .Parent Is iTextRange2.Parent Then
        '    code = code & Space(indent) & "'TODO Parent" & Chr(13)
        'End If
        'If .Start <> iTextRange2.Start Then
        '    code = code & Space(indent) & ".Start = " & LongToVBA(.Start) & Chr(13)
        'End If
        If .Text <> iTextRange2.Text Then
            code = code & Space(indent) & ".Text = " & ToVBA(.Text) & Chr(13)
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

Public Sub init(iTextRange2 As TextRange2)

    On Error GoTo err_

    With Me
        Set .Application = iTextRange2.Application
        .BoundHeight = iTextRange2.BoundHeight
        .BoundLeft = iTextRange2.BoundLeft
        .BoundTop = iTextRange2.BoundTop
        .BoundWidth = iTextRange2.BoundWidth
        .Count = iTextRange2.Count
        .Creator = iTextRange2.Creator
        Set .Font = New_iFont2(iTextRange2.Font)
        .LanguageID = iTextRange2.LanguageID
        .length = iTextRange2.length
        Set .Parent = iTextRange2.Parent
        .Start = iTextRange2.Start
        .Text = iTextRange2.Text
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
        Set .Font = New iFont2
        Call .Font.DefaultValues
    End With
End Sub



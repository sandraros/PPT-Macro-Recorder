VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "iTextRange2"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Public BoundHeight As Single
Public BoundLeft As Single
Public BoundTop As Single
Public BoundWidth As Single
Public Count As Long
Public Font As iFont2
Public LanguageID As MsoLanguageID
Public length As Long
'Public ParagraphFormat As iParagraphFormat2
Public Start As Long
Public Text As String


'ActiveWindow.Selection.ShapeRange.TextFrame2.TextRange.Start


Public Function compare(propertyName As String, iTextRange2 As iTextRange2) As Code

    Dim objCode As Code

    On Error GoTo err_

    Set objCode = New Code
    Set compare = objCode

    If IsCompared(Me) Then Exit Function

    With Me
        'If .BoundHeight <> iTextRange2.BoundHeight Then
        '    Call objCode.Add(".BoundHeight = " & SingleToVBA(.BoundHeight))
        'End If
        'If .BoundLeft <> iTextRange2.BoundLeft Then
        '    Call objCode.Add(".BoundLeft = " & SingleToVBA(.BoundLeft))
        'End If
        'If .BoundTop <> iTextRange2.BoundTop Then
        '    Call objCode.Add(".BoundTop = " & SingleToVBA(.BoundTop))
        'End If
        'If .BoundWidth <> iTextRange2.BoundWidth Then
        '    Call objCode.Add(".BoundWidth = " & SingleToVBA(.BoundWidth))
        'End If
        'If .Count <> iTextRange2.Count Then
        '    Call objCode.Add(".Count = " & LongToVBA(.Count))
        'End If
        Call objCode.AddCode(.Font.compare(".Font", iTextRange2.Font))
        If .LanguageID <> iTextRange2.LanguageID Then
            Call objCode.Add(".LanguageID = " & MsoLanguageID(.LanguageID))
        End If
        'If .length <> iTextRange2.length Then
        '    Call objCode.Add(".length = " & LongToVBA(.length))
        'End If
        'If .Start <> iTextRange2.Start Then
        '    Call objCode.Add(".Start = " & LongToVBA(.Start))
        'End If
        If .Text <> iTextRange2.Text Then
            Call objCode.Add(".Text = " & ToVBA(.Text))
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

Public Sub init(iTextRange2 As TextRange2)

    On Error GoTo err_

    With Me
        .BoundHeight = iTextRange2.BoundHeight
        .BoundLeft = iTextRange2.BoundLeft
        .BoundTop = iTextRange2.BoundTop
        .BoundWidth = iTextRange2.BoundWidth
        .Count = iTextRange2.Count
        Set .Font = New_iFont2(iTextRange2.Font)
        .LanguageID = iTextRange2.LanguageID
        .length = iTextRange2.length
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

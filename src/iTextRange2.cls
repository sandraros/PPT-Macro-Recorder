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
Public Length As Long
'Public ParagraphFormat As iParagraphFormat2
Public Parent As Object
Public Start As Long
Public Text As String

Public Runs As Collection ' Collection of iTextRange2


'ActiveWindow.Selection.ShapeRange.TextFrame2.TextRange.Start


Public Function Compare(ObjectName As String, iTextRange2 As iTextRange2) As UDiff

    Dim oCode As UDiff

    On Error GoTo err_

    Set oCode = New_UDiff(ObjectName, Me, iTextRange2)
    Set Compare = oCode

    If IsCompared(Me) Then Exit Function

    With Me
        If .Length <> 0 And iTextRange2.Length <> 0 Then
            ' A length of 0 represents the Selection object if there's just the caret inside the "text frame",
            ' i.e. no text selected, in that case Bound* properties cannot be accessed.
            If .BoundHeight <> iTextRange2.BoundHeight Then
                Call oCode.AddScalarProperty("BoundHeight", SingleToVBA(.BoundHeight))
            End If
            If .BoundLeft <> iTextRange2.BoundLeft Then
                Call oCode.AddScalarProperty("BoundLeft", SingleToVBA(.BoundLeft))
            End If
            If .BoundTop <> iTextRange2.BoundTop Then
                Call oCode.AddScalarProperty("BoundTop", SingleToVBA(.BoundTop))
            End If
            If .BoundWidth <> iTextRange2.BoundWidth Then
                Call oCode.AddScalarProperty("BoundWidth", SingleToVBA(.BoundWidth))
            End If
        End If
        'If .Count <> iTextRange2.Count Then
        '    Call oCode.AddScalarProperty("Count", LongToVBA(.Count))
        'End If
        Call oCode.AddCode(.Font.Compare("Font", iTextRange2.Font))
        If .LanguageID <> iTextRange2.LanguageID Then
            Call oCode.AddScalarProperty("LanguageID", MsoLanguageID(.LanguageID))
        End If
        'If .length <> iTextRange2.length Then
        '    Call oCode.AddScalarProperty("length", LongToVBA(.length))
        'End If
        'TODO Runs
        'If .Start <> iTextRange2.Start Then
        '    Call oCode.AddScalarProperty("Start", LongToVBA(.Start))
        'End If
        If .Text <> iTextRange2.Text Then
            Call oCode.AddScalarProperty("Text", ToVBA(.Text))
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

Public Sub init(iTextRange2 As TextRange2)

    Dim run As TextRange

    On Error GoTo err_

    With Me
        If iTextRange2.Count = 1 Then
            If iTextRange2.Length <> 0 Then
            ' A length of 0 represents the Selection object if there's just the caret inside the "text frame", no text selected.
            .BoundHeight = iTextRange2.BoundHeight
            .BoundLeft = iTextRange2.BoundLeft
            .BoundTop = iTextRange2.BoundTop
            .BoundWidth = iTextRange2.BoundWidth
            End If
        End If
        .Count = iTextRange2.Count
        Set .Font = New_iFont2(iTextRange2.Font)
        .LanguageID = iTextRange2.LanguageID
        If iTextRange2.Count = 1 Then
            .Length = iTextRange2.Length
        End If
        'TODO ParagraphFormat
        Set .Parent = iTextRange2.Parent
        Set Runs = New Collection
        If iTextRange2.Count = 1 Then
            For Each run In iTextRange2.Runs
                Call Runs.Add(New_iTextRange2(run))
            Next
            .Start = iTextRange2.Start
            .Text = iTextRange2.Text
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

    Dim defaultTextRange2 As iTextRange2

    Set defaultTextRange2 = defaultShape.TextFrame2.TextRange

    With Me
        Set .Font = defaultTextRange2.Font
        .LanguageID = defaultTextRange2.LanguageID
    End With

End Sub

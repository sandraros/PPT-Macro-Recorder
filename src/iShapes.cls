VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "iShapes"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Public Count As Long
Public Creator As Long
Public HasTitle As MsoTriState
Public Items As Collection
'Public Placeholders As iPlaceholders
Public Title As iShape


'ActivePresentation.Slides(1).Shapes.Parent


Public Function Compare(ObjectName As String, iShapes As iShapes) As UDiff

    Dim oCode As UDiff

    On Error GoTo err_

    Call goStack.Push(iShapes)

    Set oCode = New_UDiff(ObjectName, Me, iShapes)
    Set Compare = oCode

    If IsCompared(Me) Then Exit Function

    With Me
        'If .Count <> iShapes.Count Then
        '    Call oCode.AddScalarProperty("Count", LongToVBA(.Count))
        'End If
        'If .Creator <> iShapes.Creator Then
        '    Call oCode.AddScalarProperty("Creator", LongToVBA(.Creator))
        'End If
        'If .HasTitle <> iShapes.HasTitle Then
        '    Call oCode.AddScalarProperty("HasTitle", MsoTriState(.HasTitle))
        'End If
        Call oCode.AddCode(.Title.Compare("Title", iShapes.Title))
        Call CompareCollection(oCode, .Items, iShapes.Items)
    End With

    Call oCode.Wrap

    Set Compare = oCode

end_:
    Call goStack.Pop
    Exit Function

err_:
    #If DEBUG_MODE = 1 Then
        Stop
    #Else
        err.Raise err.number 'rethrows with same source and description
    #End If
    GoTo end_

End Function

Public Sub init(iShapes As Shapes)

    On Error GoTo err_

    With Me
        .Count = iShapes.Count
        .Creator = iShapes.Creator
        .HasTitle = iShapes.HasTitle
        Set .Items = New Collection
        For i = 1 To iShapes.Count
            Call .Items.Add(New_iShape(iShapes.Item(i)))
        Next
        Set .Title = New_iShape(iShapes.Title)
    End With

    Exit Sub

err_:
    #If DEBUG_MODE = 1 Then
        Stop
    #Else
        err.Raise err.number 'rethrows with same source and description
    #End If

End Sub

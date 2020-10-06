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


Public Function MR_Compare(ObjectName As String, iShapes As iShapes) As MR_Diff

    Dim oDiff As MR_Diff

    On Error GoTo err_

    'Call goStack.Push(iShapes)

    Set oDiff = New_MR_Diff(ObjectName, Me, iShapes)
    Set MR_Compare = oDiff

    If IsCompared(Me) Then Exit Function

    With Me
        'If .Count <> iShapes.Count Then
        '    Call oDiff.AddScalarProperty("Count", LongToVBA(.Count))
        'End If
        'If .Creator <> iShapes.Creator Then
        '    Call oDiff.AddScalarProperty("Creator", LongToVBA(.Creator))
        'End If
        'If .HasTitle <> iShapes.HasTitle Then
        '    Call oDiff.AddScalarProperty("HasTitle", MsoTriState(.HasTitle))
        'End If
        Call oDiff.AddDiff(.Title.MR_Compare("Title", iShapes.Title))
        Call CompareCollection(oDiff, .Items, iShapes.Items)
    End With

    Call oDiff.Wrap

    Set MR_Compare = oDiff

End_:
    'Call goStack.Pop
    Exit Function

err_:
    #If DEBUG_MODE = 1 Then
        Stop
    #Else
        err.Raise err.number 'rethrows with same source and description
    #End If
    GoTo End_

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
        If iShapes.HasTitle Then
            Set .Title = New_iShape(iShapes.Title)
        Else
            Set .Title = New_iShape()
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

    On Error GoTo err_

    With Me
    End With

    Exit Sub

err_:
    #If DEBUG_MODE = 1 Then
        Stop
    #Else
        err.Raise err.number 'rethrows with same source and description
    #End If

End Sub

Public Function OBSOLETE_AddShape(Type_ As MsoAutoShapeType, Left As Single, Top As Single, Width As Single, Height As Single) As iShape

    Dim oShape As iShape

    On Error GoTo err_

    Set oShape = New_iShape()

    With oShape
        .Type_ = Type_
        .Left = Left
        .Top = Top
        .Width = Width
        .Height = Height
        'Call .Select(Replace:=True)
    End With

    Set OBSOLETE_AddShape = oShape

    Exit Function

err_:
    #If DEBUG_MODE = 1 Then
        Stop
    #Else
        err.Raise err.number 'rethrows with same source and description
    #End If

End Function


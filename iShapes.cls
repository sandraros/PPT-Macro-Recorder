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


Public Function compare(propertyName As String, iShapes As iShapes) As Code

    Dim objCode As Code

    On Error GoTo err_

    Set objCode = New Code
    Set compare = objCode

    If IsCompared(Me) Then Exit Function

    With Me
        'If Not .Application Is iShapes.Application Then
        '    Call objCode.Add("'TODO Application")
        'End If
        'If .Count <> iShapes.Count Then
        '    Call objCode.Add(".Count = " & LongToVBA(.Count))
        'End If
        'If .Creator <> iShapes.Creator Then
        '    Call objCode.Add(".Creator = " & LongToVBA(.Creator))
        'End If
        'If .HasTitle <> iShapes.HasTitle Then
        '    Call objCode.Add(".HasTitle = " & MsoTriState(.HasTitle))
        'End If
        Call objCode.AddCode(.Title.compare(".Title", iShapes.Title))
        Call objCode.AddCode(Utility.CompareCollection(.Items, iShapes.Items))
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

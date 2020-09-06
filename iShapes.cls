VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "iShapes"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Public Application As Object
Public Count As Long
Public Creator As Long
Public HasTitle As MsoTriState
Public Items As Collection
Public Parent As Object
'Public Placeholders As iPlaceholders
Public Title As iShape


'ActivePresentation.Slides(1).Shapes.Parent


Public Function compare(propertyName As String, iShapes As iShapes, indent As Integer) As String

    On Error GoTo err_

    If IsCompared(Me) Then Exit Function

    code = ""

    With Me
        'If Not .Application Is iShapes.Application Then
        '    code = code & "'TODO Application" & Chr(13)
        'End If
        'If .Count <> iShapes.Count Then
        '    code = code & Space(indent) & ".Count = " & LongToVBA(.Count) & Chr(13)
        'End If
        'If .Creator <> iShapes.Creator Then
        '    code = code & Space(indent) & ".Creator = " & LongToVBA(.Creator) & Chr(13)
        'End If
        'If .HasTitle <> iShapes.HasTitle Then
        '    code = code & Space(indent) & ".HasTitle = " & MsoTriState(.HasTitle) & Chr(13)
        'End If
        'If Not .Parent Is iShapes.Parent Then
        '    code = code & Space(indent) & "'TODO Parent" & Chr(13)
        'End If
        code = code & .Title.compare(".Title", iShapes.Title, indent + 4)
        code = code & Utility.CompareCollection(.Items, iShapes.Items, indent + 4)
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

Public Sub init(iShapes As Shapes)

    On Error GoTo err_

    With Me
        Set .Application = iShapes.Application
        .Count = iShapes.Count
        .Creator = iShapes.Creator
        .HasTitle = iShapes.HasTitle
        Set .Items = New Collection
        For i = 1 To iShapes.Count
            Call .Items.Add(New_iShape(iShapes.Item(i)))
        Next
        Set .Parent = iShapes.Parent
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



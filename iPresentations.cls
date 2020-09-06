VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "iPresentations"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Public Application As iApplication
Public Count As Long
Public Parent As Object
Public Items As Collection


'Application.Presentations.Count


Public Function compare(propertyName As String, iPresentations As iPresentations, indent As Integer) As String

    On Error GoTo err_

    If IsCompared(Me) Then Exit Function

    code = ""

    With Me
        'If Not .Application Is iPresentations.Application Then
        '    code = code & "'TODO Application" & Chr(13)
        'End If
        If .Count <> iPresentations.Count Then
            code = code & Space(indent) & ".Count = " & LongToVBA(.Count) & Chr(13)
        End If
        If Not .Parent Is iPresentations.Parent Then
            code = code & Space(indent) & "'TODO Parent" & Chr(13)
        End If
        code = code & Utility.CompareCollection(.Items, iPresentations.Items, indent + 4)
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

Public Sub init(iPresentations As Presentations)

    On Error GoTo err_

    With Me
        Set .Application = New_iApplication(iPresentations.Application)
        .Count = iPresentations.Count
        Set .Parent = iPresentations.Parent
        Set .Items = New Collection
        For i = 1 To iPresentations.Count
            Call .Items.Add(New_iPresentation(iPresentations.Item(i)))
        Next
    End With

    Exit Sub

err_:
    #If DEBUG_MODE = 1 Then
        Stop
    #Else
        err.Raise err.number 'rethrows with same source and description
    #End If

End Sub



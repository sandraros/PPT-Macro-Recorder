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


Public Function compare(propertyName As String, iPresentations As iPresentations) As Code

    Dim objCode As Code

    On Error GoTo err_

    Set objCode = New Code
    Set compare = objCode

    If IsCompared(Me) Then Exit Function

    With Me
        'If Not .Application Is iPresentations.Application Then
        '    Call objCode.Add("'TODO Application")
        'End If
        If .Count <> iPresentations.Count Then
            Call objCode.Add(".Count = " & LongToVBA(.Count))
        End If
        'If Not .Parent Is iPresentations.Parent Then
        '    Call objCode.Add("'TODO Parent")
        'End If
        Call objCode.AddCode(Utility.CompareCollection(.Items, iPresentations.Items))
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



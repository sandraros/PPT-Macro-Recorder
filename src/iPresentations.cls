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


Public Function MR_Compare(iPresentations As iPresentations) As MR_Diff

    Dim oDiff As MR_Diff

    On Error GoTo err_

    Set oDiff = New_MR_Diff(Me, iPresentations)
    Set MR_Compare = oDiff

    If IsCompared(Me) Then Exit Function

    With Me
        'If Not .Application Is iPresentations.Application Then
        '    Call oCode.Add("'TODO Application")
        'End If
        'If .Count <> iPresentations.Count Then
        '    Call oDiff.AddScalarProperty("Count", LongToVBA(.Count))
        'End If
        'If Not .Parent Is iPresentations.Parent Then
        '    Call oCode.Add("'TODO Parent")
        'End If
        'Call CompareCollection(oDiff, .Items, iPresentations.Items)
    End With

    Call oDiff.Wrap

    Set MR_Compare = oDiff

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

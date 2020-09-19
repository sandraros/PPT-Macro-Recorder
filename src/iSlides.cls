VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "iSlides"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Public Count As Long
Public Items As Collection


'ActivePresentation.Slides.Parent


Public Function compare(propertyName As String, iSlides As iSlides) As Code

    Dim objCode As Code

    On Error GoTo err_

    Set objCode = New Code
    Set compare = objCode

    If IsCompared(Me) Then Exit Function

    With Me
        'If .Count <> iSlides.Count Then
        '    Call objCode.Add(".Count = " & LongToVBA(.Count))
        'End If
        Call objCode.AddCode(Utility.CompareCollection(.Items, iSlides.Items))
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

Public Sub init(iSlides As Slides)

    On Error GoTo err_

    With Me
        .Count = iSlides.Count
        Set .Items = New Collection
        For i = 1 To iSlides.Count
            Call .Items.Add(New_iSlide(iSlides.Item(i)))
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

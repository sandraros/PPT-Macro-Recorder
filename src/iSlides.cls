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


Public Function Compare(ObjectName As String, iSlides As iSlides) As UDiff

    Dim oCode As UDiff

    On Error GoTo err_

    Set oCode = New_UDiff(ObjectName, Me, iSlides)
    Set Compare = oCode

    If IsCompared(Me) Then Exit Function

    With Me
        'If .Count <> iSlides.Count Then
        '    Call oCode.AddScalarProperty("Count", LongToVBA(.Count))
        'End If
        Call CompareCollection(oCode, .Items, iSlides.Items)
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

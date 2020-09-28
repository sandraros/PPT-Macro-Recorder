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


Public Function MR_Compare(ObjectName As String, iSlides As iSlides) As MR_Diff

    Dim oDiff As MR_Diff

    On Error GoTo err_

    Set oDiff = New_MR_Diff(ObjectName, Me, iSlides)
    Set MR_Compare = oDiff

    If IsCompared(Me) Then Exit Function

    With Me
        'If .Count <> iSlides.Count Then
        '    Call oDiff.AddScalarProperty("Count", LongToVBA(.Count))
        'End If
        Call CompareCollection(oDiff, .Items, iSlides.Items)
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

Function AddSlide(Index As Long, pCustomLayout As iCustomLayout) As iSlide

    Dim oSlide As iSlide

    On Error GoTo err_

    Set oSlide = New_iSlide()

    With oSlide

    End With

    Set AddSlide = oSlide

    Exit Function

err_:
    #If DEBUG_MODE = 1 Then
        Stop
    #Else
        err.Raise err.number 'rethrows with same source and description
    #End If

End Function


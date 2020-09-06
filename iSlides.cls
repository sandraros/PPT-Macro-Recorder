VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "iSlides"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Public Application As iApplication
Public Count As Long
Public Parent As Object
Public Items As Collection


'ActivePresentation.Slides.Parent


Public Function compare(propertyName As String, iSlides As iSlides, indent As Integer) As String

    On Error GoTo err_

    If IsCompared(Me) Then Exit Function

    code = ""

    With Me
        'If Not .Application Is iSlides.Application Then
        '    code = code & "'TODO Application" & Chr(13)
        'End If
        'If .Count <> iSlides.Count Then
        '    code = code & Space(indent) & ".Count = " & LongToVBA(.Count) & Chr(13)
        'End If
        'If Not .Parent Is iSlides.Parent Then
        '    code = code & Space(indent) & "'TODO Parent" & Chr(13)
        'End If
        code = code & Utility.CompareCollection(.Items, iSlides.Items, indent + 4)
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

Public Sub init(iSlides As Slides)

    On Error GoTo err_

    With Me
        Set .Application = New_iApplication(iSlides.Application)
        .Count = iSlides.Count
        Set .Items = New Collection
        For i = 1 To iSlides.Count
            Call .Items.Add(New_iSlide(iSlides.Item(i)))
        Next
        Set .Parent = iSlides.Parent
    End With

    Exit Sub

err_:
    #If DEBUG_MODE = 1 Then
        Stop
    #Else
        err.Raise err.number 'rethrows with same source and description
    #End If

End Sub



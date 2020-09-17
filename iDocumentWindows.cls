VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "iDocumentWindows"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Public Application As iApplication
Public Count As Long
Public Parent As Object


'ActivePresentation.Windows.Parent


Public Function compare(propertyName As String, iDocumentWindows As iDocumentWindows) As Code

    Dim objCode As Code

    On Error GoTo err_

    Set objCode = New Code
    Set compare = objCode

    If IsCompared(Me) Then Exit Function

    With Me
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

Public Sub init(iDocumentWindows As DocumentWindows)

    On Error GoTo err_

    With Me
        Set .Application = New_iApplication(iDocumentWindows.Application)
        .Count = iDocumentWindows.Count
        Set .Parent = iDocumentWindows.Parent
    End With

    Exit Sub

err_:
    #If DEBUG_MODE = 1 Then
        Stop
    #Else
        err.Raise err.number 'rethrows with same source and description
    #End If

End Sub

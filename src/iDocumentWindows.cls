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


Public Function Compare(ObjectName As String, iDocumentWindows As iDocumentWindows) As UDiff

    Dim oCode As UDiff

    On Error GoTo err_

    Set oCode = New_UDiff(ObjectName, Me, iDocumentWindows)
    Set Compare = oCode

    If IsCompared(Me) Then Exit Function

    With Me
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

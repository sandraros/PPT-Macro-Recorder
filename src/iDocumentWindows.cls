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
Private Items As Collection


'ActivePresentation.Windows.Parent


Public Function MR_Compare(ObjectName As String, iDocumentWindows As iDocumentWindows) As MR_Diff

    Dim oDiff As MR_Diff

    On Error GoTo err_

    Set oDiff = New_MR_Diff(ObjectName, Me, iDocumentWindows)
    Set MR_Compare = oDiff

    If IsCompared(Me) Then Exit Function

    With Me
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

Public Sub init(iDocumentWindows As DocumentWindows)

    Dim oPptDocumentWindow As DocumentWindow

    On Error GoTo err_

    With Me
        Set .Application = New_iApplication(iDocumentWindows.Application)
        .Count = iDocumentWindows.Count
        Set .Parent = iDocumentWindows.Parent
        Set Items = New Collection
        For Each oPptDocumentWindow In iDocumentWindows
            Call Items.Add(GetMRObject(oPptDocumentWindow))
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

Function Item(Index As Long) As iDocumentWindow

    Set Item = Me.Items(Index)

End Function

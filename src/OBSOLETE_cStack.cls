VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "OBSOLETE_cStack"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Private oCollection As New Collection

Public Sub Push(ioAnyObject As Object)

    Call oCollection.Add(ioAnyObject)

End Sub

Public Function Pop() As Object

    If oCollection.Count > 0 Then
        Set Pop = oCollection(oCollection.Count)
        Call oCollection.Remove(oCollection.Count)
    End If

End Function

Public Property Get Top() As Object

    Set Top = Item(oCollection.Count)

End Property

Public Property Get Count() As Long

    Count = oCollection.Count

End Property

Public Property Get Item(Index As Long) As Object

    If Index >= 1 And Index <= oCollection.Count Then
        Set Item = oCollection(Index)
    End If

End Property


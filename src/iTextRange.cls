VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "iTextRange"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'SUPERSEDED WITH ITEXTRANGE2

'Public ActionSettings As iActionSettings
Public BoundHeight As Single
Public BoundLeft As Single
Public BoundTop As Single
Public BoundWidth As Single
Public Count As Long
Public Font As iFont
Public IndentLevel As Long
Public LanguageID As MsoLanguageID
Public length As Long
'Public ParagraphFormat As iParagraphFormat
Public Start As Long
Public Text As String


'ActiveWindow.Selection.ShapeRange.TextFrame.TextRange.Start


Public Function compare(propertyName As String, iTextRange As iTextRange) As Code

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

Public Sub init(iTextRange As TextRange)

    On Error GoTo err_

    With Me
        .BoundHeight = iTextRange.BoundHeight
        .BoundLeft = iTextRange.BoundLeft
        .BoundTop = iTextRange.BoundTop
        .BoundWidth = iTextRange.BoundWidth
        .Count = iTextRange.Count
        Set .Font = New_iFont(iTextRange.Font)
        .IndentLevel = iTextRange.IndentLevel
        .LanguageID = iTextRange.LanguageID
        .length = iTextRange.length
        .Start = iTextRange.Start
        .Text = iTextRange.Text
    End With

    Exit Sub

err_:
    #If DEBUG_MODE = 1 Then
        Stop
    #Else
        err.Raise err.number 'rethrows with same source and description
    #End If

End Sub

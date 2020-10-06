VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "iTextRange"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'OBSOLETE - TEXTRANGE SUPERSEDED WITH TEXTRANGE2

'Public ActionSettings As iActionSettings
Public BoundHeight As Single
Public BoundLeft As Single
Public BoundTop As Single
Public BoundWidth As Single
Public Count As Long
Public Font As iFont
Public IndentLevel As Long
Public LanguageID As MsoLanguageID
Public Length As Long
'Public ParagraphFormat As iParagraphFormat
Public Start As Long
Public Text As String

Private Runs As Collection ' Collection of iTextRange


'ActiveWindow.Selection.ShapeRange.TextFrame.TextRange.Start


Public Function MR_Compare(ObjectName As String, iTextRange As iTextRange) As MR_Diff

    Dim oDiff As MR_Diff

    On Error GoTo err_

    Set oDiff = New_MR_Diff(ObjectName, Me, iTextRange)
    Set MR_Compare = oDiff

    If IsCompared(Me) Then Exit Function

    With Me
    ' TODO
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

Public Sub init(iTextRange As TextRange)

    Dim run As TextRange

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
        If iTextRange.Runs.Count = 1 Then
            .Length = iTextRange.Length
        End If
        'TODO ParagraphFormat
        Set Runs = New Collection
        For Each run In iTextRange.Runs
            Call Runs.Add(New_iTextRange(run))
        Next
        If iTextRange.Runs.Count = 1 Then
            .Start = iTextRange.Start
            .Text = iTextRange.Text
        End If
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

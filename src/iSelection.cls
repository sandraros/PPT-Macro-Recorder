VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "iSelection"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'Public Application As iApplication
Public ChildShapeRange As iShapeRange
Public HasChildShapeRange As Boolean
'Public Parent As Object
Public shapeRange As iShapeRange
Public SlideRange As iSlideRange
'Public TextRange As iTextRange
Public TextRange2 As iTextRange2
Public Type_ As PpSelectionType ' Type


'ActiveWindow.Selection.Parent


Public Function MR_Compare(ObjectName As String, iSelection As iSelection) As MR_Diff

    Dim oDiff As MR_Diff

    On Error GoTo err_

    Set oDiff = New_MR_Diff(ObjectName, Me, iSelection)
    Set MR_Compare = oDiff

    If IsCompared(Me) Then Exit Function

    With Me
        If iSelection.HasChildShapeRange = True Then
            'Set .ChildShapeRange = New_iShapeRange(iSelection.ChildShapeRange)
        Else
            'Set .ChildShapeRange = New iShapeRange
        End If
        '.HasChildShapeRange = iSelection.HasChildShapeRange
        'Set .Parent = iSelection.Parent
        'If .Type_ = iSelection.Type_ Then
        Select Case .Type_
            Case ppSelectionShapes:
                Call oDiff.AddDiff(.shapeRange.MR_Compare("ShapeRange", iSelection.shapeRange))
            Case ppSelectionSlides:
                'TODO Set .SlideRange = New_iSlideRange(iSelection.SlideRange)
            Case ppSelectionText:
                Call oDiff.AddDiff(.TextRange2.MR_Compare("TextRange2", iSelection.TextRange2))
        End Select
        'If .Type_ <> iSelection.Type_ And .Type_ = ppSelectionNone Then
        '    Call oCode.Add("Call .Unselect")
        'End If
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

Public Sub init(iSelection As Selection)

    On Error GoTo err_

    With Me
        'Set .Application = New_iApplication(iSelection.Application)
        If iSelection.HasChildShapeRange = True Then
            Set .ChildShapeRange = New_iShapeRange(iSelection.ChildShapeRange)
        Else
            Set .ChildShapeRange = New iShapeRange
        End If
        .HasChildShapeRange = iSelection.HasChildShapeRange
        'Set .Parent = iSelection.Parent
        If iSelection.Type = ppSelectionShapes Then
            Set .shapeRange = New_iShapeRange(iSelection.shapeRange)
        Else
            Set .shapeRange = New iShapeRange
        End If
        If iSelection.Type = ppSelectionSlides Then
            Set .SlideRange = New_iSlideRange(iSelection.SlideRange)
        Else
            Set .SlideRange = New iSlideRange
        End If
        If iSelection.Type = ppSelectionText Then
            Set .TextRange2 = New_iTextRange2(iSelection.TextRange2)
        Else
            Set .TextRange2 = New iTextRange2
        End If
        .Type_ = iSelection.Type
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

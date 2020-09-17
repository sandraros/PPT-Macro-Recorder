VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "iSelection"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Public Application As iApplication
Public ChildShapeRange As iShapeRange
Public HasChildShapeRange As Boolean
Public Parent As Object
Public ShapeRange As iShapeRange
Public SlideRange As iSlideRange
'Public TextRange As iTextRange
Public TextRange2 As iTextRange2
Public iType As PpSelectionType ' Type


'ActiveWindow.Selection.Parent


Public Function compare(propertyName As String, iSelection As iSelection) As Code

    Dim objCode As Code

    On Error GoTo err_

    Set objCode = New Code
    Set compare = objCode

    If IsCompared(Me) Then Exit Function

    With Me
        If iSelection.HasChildShapeRange = True Then
            'Set .ChildShapeRange = New_iShapeRange(iSelection.ChildShapeRange)
        Else
            'Set .ChildShapeRange = New iShapeRange
        End If
        '.HasChildShapeRange = iSelection.HasChildShapeRange
        'Set .Parent = iSelection.Parent
        'If .iType = iSelection.iType Then
        Select Case .iType
            Case ppSelectionShapes:
                Call objCode.AddCode(.ShapeRange.compare(".ShapeRange", iSelection.ShapeRange))
            Case ppSelectionSlides:
                'Set .SlideRange = New_iSlideRange(iSelection.SlideRange)
            Case ppSelectionText:
                Call objCode.AddCode(.TextRange2.compare(".TextRange2", iSelection.TextRange2))
        End Select
        'End If
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

Public Sub init(iSelection As Selection)

    On Error GoTo err_

    With Me
        Set .Application = New_iApplication(iSelection.Application)
        If iSelection.HasChildShapeRange = True Then
            Set .ChildShapeRange = New_iShapeRange(iSelection.ChildShapeRange)
        Else
            Set .ChildShapeRange = New iShapeRange
        End If
        .HasChildShapeRange = iSelection.HasChildShapeRange
        Set .Parent = iSelection.Parent
        If iSelection.Type = ppSelectionShapes Then
            Set .ShapeRange = New_iShapeRange(iSelection.ShapeRange)
        Else
            Set .ShapeRange = New iShapeRange
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
        .iType = iSelection.Type
    End With

    Exit Sub

err_:
    #If DEBUG_MODE = 1 Then
        Stop
    #Else
        err.Raise err.number 'rethrows with same source and description
    #End If

End Sub



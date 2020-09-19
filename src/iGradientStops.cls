VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "iGradientStops"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Public Application As Object
Public Count As Long
Public Creator As Long
Public Items As Collection


'ActiveWindow.Selection.ShapeRange.Fill.GradientStops


Public Function compare(propertyName As String, iGradientStops As iGradientStops) As Code

    Dim objCode As Code

    On Error GoTo err_

    Set objCode = New Code
    Set compare = objCode

    If IsCompared(Me) Then Exit Function

    With Me
        'If Not .Application Is iGradientStops.Application Then
        '    Call objCode.Add("'TODO Application")
        'End If
        'If .Count <> iGradientStops.Count Then
        '    Call objCode.Add(".Count = " & LongToVBA(.Count))
        'End If
        'If .Creator <> iGradientStops.Creator Then
        '    Call objCode.Add(".Creator = " & LongToVBA(.Creator))
        'End If
        'If Not .Parent Is iGradientStops.Parent Then
        '    Call objCode.Add("'TODO Parent")
        'End If
        code2 = Utility.CompareCollection(.Items, iGradientStops.Items)
        If code2 <> "" Then
            Dim defaultGradientStops As iGradientStops
            Set defaultGradientStops = New iGradientStops
            Call defaultGradientStops.DefaultValues
            Call objCode.Add("While .Count > 0")
            Call objCode.Add("    Call .Delete(1)")
            Call objCode.Add("Wend")
            Call objCode.AddCode(Utility.CompareCollection(.Items, defaultGradientStops.Items))
        End If
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

Public Sub init(iGradientStops As Office.GradientStops)

    On Error GoTo err_

    With Me
        Set .Application = iGradientStops.Application
        .Count = iGradientStops.Count
        .Creator = iGradientStops.Creator
        Set .Items = New Collection
        For i = 1 To iGradientStops.Count
            Call .Items.Add(New_iGradientStop(iGradientStops.Item(i)))
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
    
    With Me
        Set .Items = New Collection
    End With

End Sub

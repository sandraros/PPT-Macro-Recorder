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


Public Function compare(propertyName As String, iGradientStops As iGradientStops) As String

    On Error GoTo err_

    If IsCompared(Me) Then Exit Function

    Dim code As String, code2 As String
    code = ""

    With Me
        'If Not .Application Is iGradientStops.Application Then
        '    Call addCode(code, "'TODO Application")
        'End If
        'If .Count <> iGradientStops.Count Then
        '    Call addCode(code, ".Count = " & LongToVBA(.Count))
        'End If
        'If .Creator <> iGradientStops.Creator Then
        '    Call addCode(code, ".Creator = " & LongToVBA(.Creator))
        'End If
        'If Not .Parent Is iGradientStops.Parent Then
        '    Call addCode(code, "'TODO Parent")
        'End If
        code2 = Utility.CompareCollection(.Items, iGradientStops.Items)
        If code2 <> "" Then
            Dim defaultGradientStops As iGradientStops
            Set defaultGradientStops = New iGradientStops
            Call defaultGradientStops.DefaultValues
            Call addCode(code, "While .Count > 0")
            Call addCode(code, "    Call .Delete(1)")
            Call addCode(code, "Wend")
            Call addCode(code, Utility.CompareCollection(.Items, defaultGradientStops.Items))
        End If
    End With

    Call Utility.WrapCode(propertyName, code)

    compare = code

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

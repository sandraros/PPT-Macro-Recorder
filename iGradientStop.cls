VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "iGradientStop"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Public Application As Object
Public Color As oColorFormat
Public Creator As Long
Public Position As Single
Public Transparency As Single


'ActiveWindow.Selection.ShapeRange.Fill.GradientStops(1)


Public Function compare(propertyName As String, iGradientStop As iGradientStop) As String

    On Error GoTo err_

    If IsCompared(Me) Then Exit Function

    Dim code As String
    code = ""

    With Me
        'If Not .Application Is iGradientStop.Application Then
        '    Call addCode(code, "'TODO Application")
        'End If
        Call addCode(code, .Color.compare(".Color", iGradientStop.Color))
        'If .Creator <> iGradientStop.Creator Then
        '    Call addCode(code, ".Creator = " & LongToVBA(.Creator))
        'End If
        'If Not .Parent Is iGradientStop.Parent Then
        '    Call addCode(code, "'TODO Parent")
        'End If
        If .Position <> iGradientStop.Position Then
            Call addCode(code, ".Position = " & ToVBA(.Position))
        End If
        If .Transparency <> iGradientStop.Transparency Then
            Call addCode(code, ".Transparency = " & ToVBA(.Transparency))
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

Public Sub init(iGradientStop As Office.GradientStop)

    On Error GoTo err_

    With Me
        Set .Application = iGradientStop.Application
        Set .Color = New_oColorFormat(iGradientStop.Color)
        .Creator = iGradientStop.Creator
        .Position = iGradientStop.Position
        .Transparency = iGradientStop.Transparency
    End With

    Exit Sub

err_:
    #If DEBUG_MODE = 1 Then
        Stop
    #Else
        err.Raise err.number 'rethrows with same source and description
    #End If

End Sub

Public Function create(indent) As String

    code = ""

    With Me
        Call addCode(code, Space(indent) _
                    & "Call .Insert2(" _
                    & " RGB:=" & MsoRGBTypeToVBA(.Color.RGB) & "," _
                    & " Position:=" & ToVBA(.Position) & "," _
                    & " Transparency:=" & ToVBA(.Transparency) & "," _
                    & " Index:=-1," _
                    & " Brightness:=" & ToVBA(.Color.Brightness) & ")")
    End With

    create = code

End Function

Public Function delete(indent) As String

    code = ""

    With Me
        Call addCode(code, Space(indent) & "Call .Delete(1)")
    End With

    delete = code

End Function

Sub DefaultValues()
    
    With Me
        Set .Color = New oColorFormat
        Call .Color.DefaultValues
    End With

End Sub

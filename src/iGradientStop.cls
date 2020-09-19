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


Public Function compare(propertyName As String, iGradientStop As iGradientStop) As Code

    Dim objCode As Code

    On Error GoTo err_

    Set objCode = New Code
    Set compare = objCode

    If IsCompared(Me) Then Exit Function

    With Me
        'If Not .Application Is iGradientStop.Application Then
        '    Call objCode.Add("'TODO Application")
        'End If
        Call objCode.AddCode(.Color.compare(".Color", iGradientStop.Color))
        'If .Creator <> iGradientStop.Creator Then
        '    Call objCode.Add(".Creator = " & LongToVBA(.Creator))
        'End If
        'If Not .Parent Is iGradientStop.Parent Then
        '    Call objCode.Add("'TODO Parent")
        'End If
        If .Position <> iGradientStop.Position Then
            Call objCode.Add(".Position = " & ToVBA(.Position))
        End If
        If .Transparency <> iGradientStop.Transparency Then
            Call objCode.Add(".Transparency = " & ToVBA(.Transparency))
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

Public Function create(indent) As Code

    Set objCode = New Code

    With Me
        Call objCode.Add(Space(indent) _
                    & "Call .Insert2(" _
                    & " RGB:=" & MsoRGBTypeToVBA(.Color.RGB) & "," _
                    & " Position:=" & ToVBA(.Position) & "," _
                    & " Transparency:=" & ToVBA(.Transparency) & "," _
                    & " Index:=-1," _
                    & " Brightness:=" & ToVBA(.Color.Brightness) & ")")
    End With

    Set create = objCode

End Function

Public Function delete(indent) As String

    Set objCode = New Code

    With Me
        Call objCode.Add(Space(indent) & "Call .Delete(1)")
    End With

    delete = Code

End Function

Sub DefaultValues()
    
    With Me
        Set .Color = New oColorFormat
        Call .Color.DefaultValues
    End With

End Sub

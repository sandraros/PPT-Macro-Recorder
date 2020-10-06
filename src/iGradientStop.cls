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
Public position As Single
Public Transparency As Single


'ActiveWindow.Selection.ShapeRange.Fill.GradientStops(1)


Public Function MR_Compare(ObjectName As String, iGradientStop As iGradientStop) As MR_Diff

    Dim oDiff As MR_Diff

    On Error GoTo err_

    Set oDiff = New_MR_Diff(ObjectName, Me, iGradientStop)
    Set MR_Compare = oDiff

    If IsCompared(Me) Then Exit Function

    With Me
        'If Not .Application Is iGradientStop.Application Then
        '    Call oCode.Add("'TODO Application")
        'End If
        Call oDiff.AddDiff(.Color.MR_Compare("Color", iGradientStop.Color))
        'If .Creator <> iGradientStop.Creator Then
        '    Call oDiff.AddScalarProperty("Creator", LongToVBA(.Creator))
        'End If
        'If Not .Parent Is iGradientStop.Parent Then
        '    Call oCode.Add("'TODO Parent")
        'End If
        If .position <> iGradientStop.position Then
            Call oDiff.AddScalarProperty("Position", ToVBA(.position))
        End If
        If .Transparency <> iGradientStop.Transparency Then
            Call oDiff.AddScalarProperty("Transparency", ToVBA(.Transparency))
        End If
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

Public Sub init(iGradientStop As Office.GradientStop)

    On Error GoTo err_

    With Me
        Set .Application = iGradientStop.Application
        Set .Color = New_oColorFormat(iGradientStop.Color)
        .Creator = iGradientStop.Creator
        .position = iGradientStop.position
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

Public Function create() As MR_Diff

    Stop
    'Set oDiff = New_MR_Diff(ObjectName, Me, )

    With Me
        Call oCode.Add(Space(indent) _
                    & "Call .Insert2(" _
                    & " RGB:=" & MsoRGBTypeToVBA(.Color.RGB) & "," _
                    & " Position:=" & ToVBA(.position) & "," _
                    & " Transparency:=" & ToVBA(.Transparency) & "," _
                    & " Index:=-1," _
                    & " Brightness:=" & ToVBA(.Color.Brightness) & ")")
    End With

    Set create = oCode

End Function

Public Function Delete() As String

    Stop
    'Set oDiff = New_MR_Diff(ObjectName, Me, )

    'With Me
    '    Call oCode.Add(Space(indent) & "Call .Delete(1)")
    'End With

    'Delete = Code

End Function

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

VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "iColorFormat"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'PowerPoint.ColorFormat

Public Brightness As Single
Public ObjectThemeColor As MsoThemeColorIndex
Public RGB As MsoRGBType
' Is SchemeColor obsolete?
Public SchemeColor As PpColorSchemeIndex
Public TintAndShade As Single
Public iType As MsoColorType ' Type


'ActiveWindow.Selection.ShapeRange.TextFrame2.TextRange.Font.Fill.ForeColor


Public Function Compare(ObjectName As String, iColorFormat As iColorFormat) As UDiff

    Dim oCode As UDiff

    On Error GoTo err_

    Set oCode = New_UDiff(ObjectName, Me, iColorFormat)
    Set Compare = oCode

    If IsCompared(Me) Then Exit Function

    If iColorFormat.iType = msoColorTypeMixed Then Exit Function

    With Me
        If .iType = msoColorTypeScheme _
            And .ObjectThemeColor <> iColorFormat.ObjectThemeColor Then
            ' set x = goStack.Push(iColorFormat,goStack.Top)
            ' Call oCode.Add(iColorFormat.xxxx.Compare())
            ' oCollection.Item(oCollection.Count).SelectionShapeRange =
            ' If stack contains a Shape object and does not contain iSelection.iShapeRange, generate
            ' code only if iSelection.iShapeRangeFill.ForeColor has different value for property ObjectThemeColor.
            If stopSnapShot.iSelection.iType = ppSelectionShapes Then
                Set obj1 = GetPptObject(stopSnapShot, stopSnapShot.iSelection.shapeRange.Fill.ForeColor)
                Set obj2 = GetPptObject(stopSnapShot, Me)
                If obj1 Is obj2 Then
                    a = 1
                End If
            End If
            Set oCode = New_UDiff(ObjectName, Me, iColorFormat)
            Call oCode.AddScalarProperty("ObjectThemeColor", MsoThemeColorIndex(.ObjectThemeColor))
        End If
        If .iType = msoColorTypeRGB _
            And (.iType <> iColorFormat.iType Or .RGB <> iColorFormat.RGB) Then
            Call oCode.AddScalarProperty("RGB", MsoRGBTypeToVBA(.RGB))
        End If
        'If .iType = msoColorTypeScheme And _
        '    (.iType <> iColorFormat.iType Or .SchemeColor <> iColorFormat.SchemeColor) Then
        '    Call oCode.AddScalarProperty("SchemeColor", PpColorSchemeIndex(.SchemeColor))
        'End If
        ' BRIGHTNESS must be applied after ObjectThemeColor which resets it
        If .Brightness <> iColorFormat.Brightness Then
            Call oCode.AddScalarProperty("Brightness", SingleToVBA(.Brightness))
        End If
        If .TintAndShade <> iColorFormat.TintAndShade Then
            Call oCode.AddScalarProperty("TintAndShade", SingleToVBA(.TintAndShade))
        End If
        'If .iType <> iColorFormat.iType Then
        '    Call oCode.Add("'.Type", MsoColorType(.iType))
        'End If
    End With

    Call oCode.Wrap

    Set Compare = oCode

    Exit Function

err_:
    #If DEBUG_MODE = 1 Then
        Stop
    #Else
        err.Raise err.number 'rethrows with same source and description
    #End If

End Function

Public Sub init(iColorFormat As PowerPoint.ColorFormat)

    On Error GoTo err_

    If iColorFormat.Type = msoColorTypeMixed Then Exit Sub

    With Me
        .Brightness = iColorFormat.Brightness
        If iColorFormat.Type = msoColorTypeScheme Then
            .ObjectThemeColor = iColorFormat.ObjectThemeColor
        End If
        'If iColorFormat.Type = msoColorTypeRGB Then
        'RGB should be returned for at least Type = msoColorTypeRGB and msoColorTypeScheme
        '(RGB is calculated for msoColorTypeScheme, based on ObjectThemeColor and TintAndShade)
        .RGB = iColorFormat.RGB
        'End If
        'If iColorFormat.Type = msoColorTypeScheme Then
        '    .SchemeColor = iColorFormat.SchemeColor
        'End If
        .TintAndShade = iColorFormat.TintAndShade
        .iType = iColorFormat.Type
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
    End With
End Sub

